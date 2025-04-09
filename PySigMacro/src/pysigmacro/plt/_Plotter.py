#!/usr/bin/env python3
# -*- coding: utf-8 -*-
# Timestamp: "2025-04-09 20:39:18 (ywatanabe)"
# File: /home/ywatanabe/win/documents/SigMacro/PySigMacro/src/pysigmacro/plt/_Plotter.py
# ----------------------------------------
import os
__FILE__ = (
    "/home/ywatanabe/win/documents/SigMacro/PySigMacro/src/pysigmacro/plt/_Plotter.py"
)
__DIR__ = os.path.dirname(__FILE__)
# ----------------------------------------

# src/pysigmacro/plt/_plotter.py
import pandas as pd
from ..const import COLORS, BGRA, PLOT_TYPES
from ..data import create_padded_df, create_graph_wizard_params
from ..demo import gen_visual_params, update_visual_params_with_nice_ticks
from ..demo import JNBGenerator


class Plotter:
    """
    A class to facilitate creating SigmaPlot figures with multiple plot types
    in a style similar to matplotlib.
    """

    def __init__(self):
        """Initialize the Plotter."""
        # List to store individual plot requests
        self.plot_requests = []
        # Dictionary to store user-specified visual parameters
        self.visual_params = {}

    def add(self, plot_type, data_dict):
        """
        Add a plot layer to the figure.

        Args:
            plot_type (str): The type of plot (e.g., 'line', 'scatter', 'bar').
                             Must be one of the supported plot types.
            data_dict (dict): DataFrame containing the data for this plot.
                              Columns should match the expected format for
                              the chosen plot_type (e.g., 'x', 'y', 'yerr').

        Raises:
            ValueError: If plot_type is not supported.
        """

        data_df = create_padded_df(data_dict)

        if plot_type not in PLOT_TYPES:
            raise ValueError(
                f"Unsupported plot_type: {plot_type}. Supported types: {PLOT_TYPES}"
            )

        # Store the request
        self.plot_requests.append({"type": plot_type, "data": data_df.copy()})

    def set_params(self, **kwargs):
        """
        Set or override visual parameters for the plot (e.g., labels, limits).

        Args:
            **kwargs: Keyword arguments matching visual parameter keys
                      (e.g., xlabel="Time", xmin=0, ymax=100).
        """
        self.visual_params.update(kwargs)

    def _format_user_data(self, plot_type, user_df, i_plot):
        """
        Formats the user-provided DataFrame for a specific plot type.
        Adds default BGRA color if not present.
        Renames columns to standard names expected by internal functions if needed.
        (This is a basic implementation and might need refinement per plot type)
        """
        formatted_df = user_df.copy()

        # Add default BGRA if not present
        if "bgra" not in formatted_df.columns:
            # Select color based on plot index
            bgra_color = BGRA[COLORS[i_plot % len(COLORS)]].copy()
            # Apply this BGRA to all rows (needs padding logic)
            # For simplicity, just add it as a potential column; create_padded_df handles length
            # A more robust approach would apply it correctly if length matches
            formatted_df["bgra"] = [bgra_color] * len(
                formatted_df
            )  # Basic way, padding handles mismatch

        # --- TODO: Add column renaming logic based on plot_type if necessary ---
        # Example: if plot_type == 'line' and 'time' in formatted_df.columns:
        #              formatted_df = formatted_df.rename(columns={'time': 'x'})
        # Example: Ensure 'x', 'y' exist for scatter/line

        formatted_df.columns = [f"{col} {i_plot}" for col in formatted_df.columns]

        return formatted_df

    def _merge_visual_params(self, default_df, user_dict):
        """Merges default visual parameters with user overrides."""
        # Convert default DataFrame to dict
        default_dict = default_df.set_index(default_df.columns[0])[
            default_df.columns[1]
        ].to_dict()

        # Override defaults with user settings
        default_dict.update(user_dict)

        # Convert back to the required structure (list of key-value pairs for DataFrame)
        # This part needs careful handling to match the structure expected by create_padded_df
        # Assuming create_padded_df can handle a flat dictionary for visual params:
        return default_dict  # Return the merged dictionary

    def render(self, output_dir=".", filename_base="plot", keep_orig_figures=False, **kwargs):
        """
        Generates the SigmaPlot figure based on the added plot requests.

        Args:
            output_dir (str): Directory to save the generated .jnb and image files.
            filename_base (str): Base name for the output files.
            **kwargs: Additional visual parameters to set for this render call.

        Returns:
            str: Path to the generated JNB file.

        Raises:
            ValueError: If no plots have been added.
            Exception: Propagates Exceptions from JNBGenerator.
        """
        if not self.plot_requests:
            raise ValueError(
                "No plots added. Use the add() method before rendering."
            )

        # Update visual params with any passed directly to render
        self.set_params(**kwargs)

        # --- 1. Prepare Combined Data DataFrame ---
        plot_types = [req["type"] for req in self.plot_requests]
        final_data_chunks_df = pd.DataFrame()

        for i_plot, request in enumerate(self.plot_requests):
            plot_type = request["type"]
            user_df = request["data"]

            # Get Graph Wizard parameters for this plot
            gw_df = create_graph_wizard_params(
                plot_type, i_plot, label=f"{plot_type} {i_plot}"
            )

            # Format user data (add BGRA, potentially rename cols)
            formatted_user_df = self._format_user_data(
                plot_type, user_df, i_plot
            )

            # Combine GW params and formatted data for this chunk
            chunk_df = create_padded_df(gw_df, formatted_user_df)

            # Concatenate horizontally with previous chunks
            final_data_chunks_df = create_padded_df(
                final_data_chunks_df, chunk_df
            )

        # --- 2. Prepare Visual Parameters ---
        # Use the first plot type to get default visual params structure
        merged_visual_params_df = gen_visual_params(plot_types[0], **self.visual_params)
        # default_visual_params_df = gen_visual_params(plot_types[0], **self.visual_params)

        # --- 3. Combine Params and Data ---
        final_df = create_padded_df(
            merged_visual_params_df, final_data_chunks_df
        )

        # --- 4. Apply Nice Ticks ---
        # Note: update_visual_params_with_nice_ticks modifies the df inplace or returns a modified one
        # Need to adapt its input/output or replicate its logic here based on merged_params_dict and final_data_chunks_df
        # For now, we pass the combined final_df, but this might need adjustment
        # It's safer to calculate ticks *before* combining into final_df if possible
        try:
            final_df = update_visual_params_with_nice_ticks(
                plot_types[0], merged_visual_params_df, final_data_chunks_df
            )
            # Re-combine after potential modification by nice_ticks
            final_df = create_padded_df(final_df, final_data_chunks_df)
        except Exception as e:
            print(f"Warning: Could not apply nice ticks - {e}")
            # Proceed with the potentially un-updated final_df

        # --- 5. Generate JNB ---
        generator = JNBGenerator(plot_types, df=final_df, from_demo_csv=False)

        # Manually set paths for the generator
        templates_dir = os.getenv(
            "SIGMACRO_TEMPLATES_DIR",
            os.path.join(os.path.dirname(__file__), "..", "templates"),
        )  # Adjust path as needed
        generator.path_jnb_template = os.path.join(
            templates_dir, "jnb", "template.JNB"
        )

        output_base = os.path.join(output_dir, filename_base)
        generator.path_jnb = f"{output_base}.jnb"
        # Optional: Set other paths if needed by generator internals or export methods
        generator.path_tif = f"{output_base}.tif"
        generator.path_gif = f"{output_base}.gif"
        # ... other formats

        # Ensure output directory exists
        os.makedirs(output_dir, exist_ok=True)

        # Run the generation process
        generated_jnb_path = generator.run(keep_orig_figures=keep_orig_figures)

        return generated_jnb_path

# EOF