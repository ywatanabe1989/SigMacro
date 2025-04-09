#!/usr/bin/env python3
# -*- coding: utf-8 -*-
# Timestamp: "2025-04-01 13:49:36 (ywatanabe)"
# File: /home/ywatanabe/win/documents/SigMacro/PySigMacro/src/pysigmacro/com/_BaseCOMWrapper.py
# ----------------------------------------
import os
__FILE__ = (
    "/home/ywatanabe/win/documents/SigMacro/PySigMacro/src/pysigmacro/com/_BaseCOMWrapper.py"
)
__DIR__ = os.path.dirname(__FILE__)
# ----------------------------------------

import re

import win32com.client

from ._base import get_wrapper
from ._inspect import inspect


class BaseCOMWrapper:
    """
    A wrapper class that exposes COM object properties, methods and child objects
    using a pythonic interface based on inspection.
    """

    __classname__ = "BaseCOMWrapper"

    def __init__(self, com_object, access_path=""):
        """
        Initialize the wrapper with a COM object.
        """
        self._com_object = com_object
        self._access_path = access_path
        self._inspected = inspect(com_object)
        self._cached_objects = {}
        self._access_path = access_path
        self._wrap_properties()

    def _wrap_properties(self):
        """
        Wrap COM object properties based on inspected data.
        """
        for _, row in self._inspected.iterrows():
            if row["Type"] in ("Property", "Object"):
                self._create_property(row["Name"], row["Type"])
        # Create dynamic class for properties
        cls_name = f"Dynamic{type(self).__name__}"
        bases = (type(self),)
        self.__class__ = type(cls_name, bases, {})

    def _create_property(self, name, typ):
        """
        Create a Python property wrapping the COM object's attribute.
        """
        if typ == "Property":

            def getter(instance):
                try:
                    return getattr(instance._com_object, name)
                except Exception:
                    return None

            def setter(instance, value):
                try:
                    if hasattr(value, "_com_object"):
                        value = value._com_object
                    setattr(instance._com_object, name, value)
                except Exception as e:
                    print(f"Warning: Could not set property {name}: {e}")

            setattr(type(self), name, property(getter, setter))
        elif typ == "Object":

            def getter(instance):
                try:
                    if name in instance._cached_objects:
                        return instance._cached_objects[name]
                    value = getattr(instance._com_object, name)
                    # Calculate new access_path for this object
                    new_access_path = (
                        f"{instance._access_path}.{name}"
                        if instance._access_path
                        else name
                    )
                    wrapped = get_wrapper(value, new_access_path, self.path)
                    if hasattr(instance, "path"):
                        wrapped.path = instance.path
                    instance._cached_objects[name] = wrapped
                    return wrapped
                except Exception:
                    return None

            def setter(instance, value):
                try:
                    if hasattr(value, "_com_object"):
                        value = value._com_object
                    setattr(instance._com_object, name, value)
                except Exception as e:
                    print(f"Warning: Could not set property {name}: {e}")

            setattr(type(self), f"{name}_obj", property(getter, setter))

    @property
    def access_path(self):
        return self._access_path

    @property
    def list(self):
        """Return a list of all items' names in a collection"""
        try:
            if hasattr(self._com_object, "Count"):
                count = self._com_object.Count
                names_list = []
                for i in range(count):
                    try:
                        name = getattr(
                            self._com_object[i], "Name", f"Item {i}"
                        )
                        names_list.append(name)
                    except:
                        names_list.append(f"(Unnamed item)")
                return names_list
            raise AttributeError("Object does not support listing")
        except Exception as e:
            print(f"Error listing items: {e}")
            return []

    def list_items(self):
        """Display all items in a collection with their indices"""
        try:
            if hasattr(self._com_object, "Count"):
                count = self._com_object.Count
                print(f"{self._access_path or 'Collection'} - {count} items:")
                names_list = self.list
                for i, name in enumerate(names_list):
                    print(f"  [{i}] {name}")
                return names_list
            raise AttributeError("Object does not support listing")
        except Exception as e:
            print(f"Error listing items: {e}")
            return []

    def find_indices(self, pattern):
        """Find items matching pattern and return their indices"""
        indices = []
        try:
            names = self.list
            if names:
                regex = re.compile(pattern, re.IGNORECASE)
                for i, name in enumerate(names):
                    if regex.search(name):
                        indices.append(i)
            return indices
        except Exception as e:
            print(f"Error finding items: {e}")
            return []

    def clean(self):
        """Base cleaning implementation"""
        print(f"No specific cleaning implemented for {self._access_path}")
        return self

    def _wrap_com_method(self, method, name):
        """Wrap a COM method to properly handle return values"""

        def wrapped_method(*args, **kwargs):
            try:
                result = method(*args, **kwargs)
                if hasattr(result, "_oleobj_"):
                    # Calculate new access_path for method result
                    new_access_path = (
                        f"{self._access_path}.{name}()"
                        if self._access_path
                        else f"{name}()"
                    )
                    wrapped = get_wrapper(result, new_access_path, self.path)
                    # Pass the path to the new wrapper
                    if hasattr(self, "_path"):
                        wrapped._path = self._path
                    return wrapped
                return result
            except Exception as e:
                print(f"Error calling method {name}: {e}")
                raise

        return wrapped_method

    def _wrap_com_object(self, com_obj, name=None):
        """Wrap a COM object with appropriate wrapper"""
        if name is None:
            name = "unknown"
        new_access_path = (
            f"{self._access_path}.{name}" if self._access_path else name
        )
        wrapped = get_wrapper(com_obj, new_access_path, self.path)
        # Pass the path to the new wrapper
        if hasattr(self, "_path") and self._path:
            wrapped._path = self._path
        return wrapped

    def __getattr__(self, name):
        """Handle attribute access for COM objects"""
        try:
            # Try to get the attribute directly from the COM object
            attr = getattr(self._com_object, name)
            # If it's a callable method
            if callable(attr):
                # Wrap the method to handle return values properly
                def method_wrapper(*args, **kwargs):
                    # Convert BaseCOMWrapper arguments to their COM objects
                    com_args = []
                    for arg in args:
                        if isinstance(arg, BaseCOMWrapper):
                            com_args.append(arg._com_object)
                        else:
                            com_args.append(arg)
                    # Call the method with the unwrapped arguments
                    result = attr(*com_args, **kwargs)
                    # If the result is a COM object, wrap it properly
                    if hasattr(result, "_oleobj_"):
                        new_access_path = (
                            f"{self._access_path}.{name}()"
                            if self._access_path
                            else f"{name}()"
                        )
                        return self._wrap_com_object(result, name)
                    return result

                return method_wrapper
            # If it's a COM object itself
            elif hasattr(attr, "_oleobj_"):
                # Wrap it in appropriate wrapper
                wrapped = self._wrap_com_object(attr, name)
                return wrapped
            # Otherwise it's a simple value property
            else:
                return attr
        except Exception as e:
            # Check if it's in our known methods list
            from ..const._SIGMAPLOT_METHODS import SIGMAPLOT_METHODS

            if name in SIGMAPLOT_METHODS:
                # Return a method that will use Invoke directly
                def invoke_method(*args, **kwargs):
                    try:
                        # Convert BaseCOMWrapper arguments to their COM objects
                        com_args = []
                        for arg in args:
                            if isinstance(arg, BaseCOMWrapper):
                                com_args.append(arg._com_object)
                            else:
                                com_args.append(arg)
                        # Get the dispatch ID
                        dispid = self._com_object._oleobj_.GetIDsOfNames(
                            0, [name]
                        )[0]
                        # Reverse the arguments for COM
                        reversed_args = list(reversed(com_args))
                        # Call the method using Invoke with DISPATCH_METHOD flag (1)
                        import pythoncom

                        result = self._com_object._oleobj_.Invoke(
                            dispid, 0, 1, reversed_args
                        )
                        # If the result is a COM object, wrap it properly
                        if hasattr(result, "_oleobj_"):
                            new_access_path = (
                                f"{self._access_path}.{name}()"
                                if self._access_path
                                else f"{name}()"
                            )
                            return self._wrap_com_object(result, name)
                        return result
                    except Exception as invoke_error:
                        print(f"Error invoking method {name}: {invoke_error}")
                        raise

                return invoke_method
            # Attribute not found on COM object
            raise AttributeError(f"Attribute {name} not found") from e

    def __dir__(self):
        """Return list of available attributes"""
        attrs = set()
        try:
            attrs.update(dir(self._com_object))
        except Exception:
            pass
        attrs.update(super().__dir__())
        for _, row in self._inspected.iterrows():
            typ = row["Type"]
            name = row["Name"]
            if typ == "Property":
                attrs.add(name)
            elif typ == "Object":
                attrs.add(f"{name}_obj")
            elif typ == "Method":
                attrs.add(name)
        return sorted(attrs)

    def __repr__(self):
        """String representation of the wrapper"""
        classname = getattr(self, "__classname__", type(self).__name__)
        name = self.name if hasattr(self, "name") else "unnamed"
        return f"<{classname} for {name} at {self.access_path}>"

    def __str__(self):
        return self.__repr__()

    def __len__(self):
        """Return Count for collections"""
        if hasattr(self._com_object, "Count"):
            return self._com_object.Count
        return 0

    def Item(self, key):
        """Access item by index, handling special case for last item"""
        if key == -1 and hasattr(self._com_object, "Count"):
            key = self._com_object.Count - 1
        result = self._com_object(key)
        access_path = (
            f"{self._access_path}({key})" if self._access_path else f"({key})"
        )
        wrapped = get_wrapper(result, access_path, self.path)
        # Pass the path to the new wrapper
        if hasattr(self, "_path"):
            wrapped._path = self._path
        return wrapped

    def __call__(self, key):
        """Make the wrapper callable for accessing items"""
        return self.Item(key)

    def __getitem__(self, key):
        """Access items using square bracket notation"""
        if key == -1 and hasattr(self._com_object, "Count"):
            key = self._com_object.Count - 1
        result = self._com_object(key)
        access_path = (
            f"{self._access_path}[{key}]" if self._access_path else f"[{key}]"
        )
        wrapped = get_wrapper(result, access_path, self.path)
        # Pass the path to the new wrapper
        if hasattr(self, "_path"):
            wrapped._path = self._path
        return wrapped

    @property
    def path(self):
        """Get the file path for this object"""
        if hasattr(self, "_path") and self._path:
            return self._path
        # For collection items, try to get parent path
        if "_" in self._access_path and not self._access_path.endswith("_"):
            parent_path = self._access_path.split("_")[0]
            if parent_path and not parent_path.startswith("SigmaPlot"):
                return parent_path
        # Transform access_path to a more file-friendly format as fallback
        return (
            self._access_path.replace(".", "_")
            .replace("(", "_")
            .replace(")", "_")
            .replace("[", "_")
            .replace("]", "_")
        )

    @path.setter
    def path(self, value):
        """Set the path for this object"""
        self._path = value
        # Propagate to cached objects
        for key, obj in self._cached_objects.items():
            if hasattr(obj, "path"):
                obj._path = value

    @staticmethod
    def get_active_document():
        """Get the currently active document"""

        app = win32com.client.Dispatch("SigmaPlot.Application").Application

        try:
            active_doc = app.ActiveDocument
            if active_doc:
                from ._base import get_wrapper

                return get_wrapper(active_doc, "ActiveDocument")
            else:
                return None
        except Exception as e:
            print(f"Error getting active document: {e}")
            return None

    def activate(self, visible=False):
        """Activate this object (make it visible and current)"""
        try:
            if visible:
                # First try to make the application visible
                if "Application" in self._access_path or self._access_path == "":
                    if hasattr(self._com_object, "Visible"):
                        self._com_object.Visible = True

            # Try different activation methods
            if hasattr(self._com_object, "Activate"):
                self._com_object.Activate()
            if hasattr(self._com_object, "SetActive"):
                self._com_object.SetActive()
            if hasattr(self._com_object, "Select"):
                self._com_object.Select()
            if hasattr(self._com_object, "SetObjectCurrent"):
                self._com_object.SetObjectCurrent()
            if hasattr(self._com_object, "Open"):
                self._com_object.Open()
            else:
                # Try to use the Window property if available
                if hasattr(self._com_object, "Window"):
                    window = self._com_object.Window
                    if hasattr(window, "Activate"):
                        window.Activate()
                else:
                    print(f"Warning: Object {self._access_path} doesn't support any known activation method")

            return self
        except Exception as e:
            print(f"Error activating object {self._access_path}: {e}")
            return self

# EOF