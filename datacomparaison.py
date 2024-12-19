import numpy as np
import pandas as pd
import matplotlib.pyplot as plt
import streamlit as st
from io import BytesIO
from scipy.interpolate import make_interp_spline

# Streamlit application configuration
st.set_page_config(
    page_title="Interactive Acoustic Analysis",
    page_icon=":chart_with_upwards_trend:",
    layout="centered",
    initial_sidebar_state="expanded"
)

# Predefined frequency list
frequencies = np.array([200, 250, 315, 400, 500, 630, 800, 1000,
    1250, 1600, 2000, 2500, 3150, 4000, 5000,
    6300, 8000, 10000])

# Function to load Excel files
def load_excel(file):
    """
    Load an Excel file, supporting both .xls and .xlsx formats.
    """
    try:
        # Detect file format
        if file.name.endswith(".xls"):
            df = pd.read_excel(file, sheet_name="Macro", engine="xlrd", header=None)
        else:
            df = pd.read_excel(file, sheet_name="Macro", engine="openpyxl", header=None)
        return df
    except Exception as e:
        st.error(f"Error reading the file: {e}")
        return None

# Function to extract data
def extract_data(df):
    """
    Extract valid series from the DataFrame.
    """
    extracted_data = []

    for col_start, name_cell, values_range in [
        (0, "A1", "C3:C20"), (4, "E1", "G3:G20"),
        (8, "I1", "K3:K20"), (12, "M1", "O3:O20"),
        (16, "Q1", "S3:S20"), (20, "U1", "W3:W20")
    ]:
        try:
            # Read the series name
            name = df.iloc[0, col_start]

            # Read the values (avoiding non-numeric errors like #DIV/0!)
            values = pd.to_numeric(
                df.iloc[2:20, col_start + 2], errors="coerce"
            ).values  # `coerce` turns errors into NaN

            # Check if at least one value is numeric
            if not np.isnan(values).all():
                extracted_data.append({
                    "name": name,  # Use only the sample name
                    "values": values
                })
        except Exception:
            continue

    return extracted_data

# File upload
uploaded_files = st.sidebar.file_uploader(
    "Upload your Excel files (.xls or .xlsx)",
    type=["xls", "xlsx"],
    accept_multiple_files=True
)

# Store all extracted series
all_series = []

# Load files and extract data
if uploaded_files:
    for file in uploaded_files:
        st.subheader(f"Extracted data from: {file.name}")

        # Load the file
        df = load_excel(file)
        if df is None:
            continue

        # Extract valid data
        extracted_data = extract_data(df)

        # Add the series to the global list
        all_series.extend(extracted_data)

# Display selection options if valid series exist
if all_series:
    # List of available series names
    series_names = [series["name"] for series in all_series]

    # Allow users to select which series to display
    selected_series_names = st.sidebar.multiselect(
        "Choose series to display",
        options=series_names,
        default=series_names  # By default, all series are selected
    )

    # Function to smooth the data with logarithmic interpolation
    def smooth_curve(frequencies, values, num_points=40):
        """
        Smooth the curve using spline interpolation on a logarithmic scale.
        
        Parameters:
        - frequencies: Original frequency values (x-axis).
        - values: Original absorption values (y-axis).
        - num_points: Number of points for the smoothed curve.
        
        Returns:
        - smoothed_freq: Smoothed frequency values.
        - smoothed_values: Smoothed absorption values.
        """
        # Remove NaN values for proper interpolation
        valid_indices = ~np.isnan(values)
        frequencies = frequencies[valid_indices]
        values = values[valid_indices]

        # Logarithmic interpolation (log scale for frequency)
        log_frequencies = np.log(frequencies)  # Take the logarithm of the frequencies
        smoothed_log_freq = np.linspace(log_frequencies.min(), log_frequencies.max(), num_points)
        
        # Cubic spline interpolation on the log scale
        spline = make_interp_spline(log_frequencies, values, k=3)
        smoothed_values = spline(smoothed_log_freq)
        
        # Convert back to the frequency scale (exponentiate the smoothed log frequencies)
        smoothed_freq = np.exp(smoothed_log_freq)
        
        return smoothed_freq, smoothed_values

    # Filter selected series
    selected_series = [series for series in all_series if series["name"] in selected_series_names]

    # Generate a single graph for all selected series
    fig, ax = plt.subplots(figsize=(12, 8))
    for series in selected_series:
        # Smooth the data
        smoothed_freq, smoothed_values = smooth_curve(frequencies, series["values"])
        
        # Plot the smoothed curve
        ax.plot(smoothed_freq, smoothed_values, label=series["name"], marker="none")
    
    # Customize the graph
    ax.set_title("Absorption Curves (Selected Series)")
    ax.set_xscale("log")
    ax.set_xticks(frequencies)
    ax.get_xaxis().set_major_formatter(plt.ScalarFormatter())
    ax.set_xlabel("Frequency (Hz)")
    ax.set_ylabel("Absorption Coefficient")
    ax.legend(loc="best", fontsize=10)
    ax.grid(True, which="both", linestyle="--", linewidth=0.5)
    
    # Display the graph in Streamlit
    st.pyplot(fig)

    # Function to save the graph as a PDF
    def save_as_pdf(fig):
        pdf_buffer = BytesIO()
        fig.savefig(pdf_buffer, format="pdf")
        pdf_buffer.seek(0)
        return pdf_buffer

    # Function to save the graph as a JPEG
    def save_as_jpeg(fig):
        jpeg_buffer = BytesIO()
        fig.savefig(jpeg_buffer, format="jpeg", dpi=300)
        jpeg_buffer.seek(0)
        return jpeg_buffer

    # Add download buttons
    st.download_button(
        label="Download the graph as PDF",
        data=save_as_pdf(fig),
        file_name="absorption_graph.pdf",
        mime="application/pdf"
    )
    st.download_button(
        label="Download the graph as JPEG",
        data=save_as_jpeg(fig),
        file_name="absorption_graph.jpeg",
        mime="image/jpeg"
    )

else:
    st.info("Please upload at least one Excel file with valid series to start the analysis.")
