# Interactive Acoustic Analysis

## Description

This project provides an interactive tool for acoustic analysis of absorption coefficients at different frequencies using data from Excel files. The app allows users to upload Excel files containing acoustic data, visualize absorption curves, and smooth them using spline interpolation. It also offers the option to download the graphs in PDF or JPEG formats.

### Key Features:
- **Upload Excel Files**: The application supports `.xls` and `.xlsx` formats.
- **Interactive Plotting**: Choose which series to display and visualize absorption curves.
- **Smoothing**: Interpolation on a logarithmic frequency scale to smooth the curves.
- **Download Graphs**: Download the resulting graphs in PDF or JPEG formats.

## How It Works

1. **Upload Files**: Upload one or more Excel files that contain acoustic data. The app will automatically extract relevant series.
2. **Select Series**: Once the data is extracted, you can select which series to visualize.
3. **Smooth Curves**: The app uses cubic spline interpolation to smooth the absorption curves. The interpolation takes place on a logarithmic frequency scale to ensure better visualization of the data, especially in higher frequencies.
4. **Download Options**: After the graph is generated, you can download it as either a PDF or a JPEG file.

## Requirements

To run the app locally, you need to have the following libraries installed:

- `numpy`
- `pandas`
- `matplotlib`
- `streamlit`
- `scipy`

You can install these dependencies by running:

```bash
pip install numpy pandas matplotlib streamlit scipy
```

## App Link

You can access the live version of the app at the following link:  
[Interactive Acoustic Analysis App](https://github.com/LinoVation1312/alpha_cab/)

## License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.

## Acknowledgements

- **Streamlit**: For creating a simple way to build interactive web apps in Python.
- **Scipy**: For providing advanced interpolation functions.
- **Matplotlib**: For visualizing the data.
