#!/usr/bin/env python3
# -*- coding: utf-8 -*-
# Timestamp: "2025-03-15 02:44:06 (ywatanabe)"
# File: /home/ywatanabe/proj/SigMacro/PySigMacro/src/pysigmacro/utils/_com_wrap.py
# ----------------------------------------
import os
__FILE__ = (
    "/home/ywatanabe/proj/SigMacro/PySigMacro/src/pysigmacro/utils/_com_wrap.py"
)
__DIR__ = os.path.dirname(__FILE__)
# ----------------------------------------

from ._sigmaplot_inspect import inspect

# import pythoncom


class COMWrapper:
    """
    A wrapper class that exposes COM object properties, methods and child objects
    using a pythonic interface based on inspection.
    """

    def __init__(self, com_object):
        """
        Initialize the wrapper with a COM object.
        """
        self._com_object = com_object
        self._inspected = inspect(com_object)
        self._wrap_properties()
        # Removed _wrap_methods call to avoid premature method invocation
        # self._wrap_methods()

    def _wrap_properties(self):
        """
        Wrap COM object properties based on inspected data.
        """
        for _, row in self._inspected.iterrows():
            if row["Type"] in ("Property", "Object"):
                self._create_property(row["Name"], row["Type"])

    def _create_property(self, name, typ):
        """
        Create a Python property wrapping the COM object's attribute.
        For 'Property' type, create a property with the given name.
        For 'Object' type, create only the property with _obj suffix.
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
                    if not hasattr(instance, "_cached_objects"):
                        instance._cached_objects = {}
                    if name in instance._cached_objects:
                        return instance._cached_objects[name]
                    value = getattr(instance._com_object, name)
                    wrapped = com_wrap(value)
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

    def _wrap_methods(self):
        """
        Wrap COM object methods based on inspected data, adding them as Python methods.
        """
        for _, row in self._inspected.iterrows():
            if row["Type"] == "Method":
                name = row["Name"]
                method = getattr(self._com_object, name, None)
                if method and callable(method):
                    setattr(self, name, method)

    def __getattr__(self, name):
        try:
            attr = getattr(self._com_object, name)
            if callable(attr):

                def method_wrapper(*args, **kwargs):
                    result = attr(*args, **kwargs)
                    if hasattr(result, "_oleobj_"):
                        return com_wrap(result)
                    return result

                return method_wrapper
            else:
                if hasattr(attr, "_oleobj_"):
                    return com_wrap(attr)
                return attr
        except Exception as e:
            raise AttributeError(f"Attribute {name} not found") from e

    def __dir__(self):
        """
        Return a list of attributes from the COM object's inspection data
        without triggering the property getters.
        """
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
        """
        Return a string representation of the COMWrapper.
        """
        try:
            name = getattr(self._com_object, "Name", "Unknown")
        except Exception:
            name = "Unknown"
        return f"<COMWrapper for {name}>"

    def __len__(self):
        return self._com_object.Count


def com_wrap(com_object):
    """
    Create a fresh Python wrapper for a COM object.
    """
    return COMWrapper(com_object)

# EOF