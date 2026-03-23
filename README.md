# Mass-Spectrometry-Data-Extraction

This Python-based software provides a streamlined, automated pipeline for extracting and aligning mass spectrometry data from both Thermo (HCD/TRMS) and Bruker (CID/TRMS) instruments.
The script ingests converted mass spectrometry files (.ms1, .ms2, or .ascii), divides the run into discrete time segments, and extracts the corresponding mass-to-charge (m/z) intensities. It then aggregates this data into a single, cleanly formatted Excel workbook.

Key Features:
Unified Processing: Handles both Thermo and Bruker data structures through a single, standardized interface.
Smart Peak Detection: Automatically identifies the precursor ion and the most abundant fragment ions. It calculates the total intensity of every m/z across the entire run to ensure it tracks true fragments, actively ignoring random, split-second electronic noise.
Strict Data Control: Allows users to set a hard limit on the number of fragment columns, ensuring the final exported data is concise and strictly relevant.
Batch Processing: Users can queue multiple data files at once. The tool processes each run sequentially and automatically aligns them vertically in the final Excel export for easy comparison.
Analysis-Ready Output: Generates a formatted Excel file with bold headers and perfectly aligned columns, drastically reducing the time spent preparing data for final analysis.

## Prerequisites

To run this script, you will need Python installed on your computer, along with a few standard data processing libraries.

1. Install Spyder or Python
2. Copy thid Python script and Run

## Data Preparation

Before running the script, ensure your mass spectrometry data is in the correct format:
* **Thermo Data:** Raw files must be converted to `.ms1` or `.ms2` format using **MSConvert**.
* **Bruker Data:** Data must be exported to `.ascii` format using the `Export_data_seg_times` method.

## How to Use the Script

1. **Run the Script:** A series of graphical windows will appear to guide you.
2. **Select Instrument:** Click the button corresponding to the instrument used to generate your data (Thermo or Bruker).
3. **Select Files:** A file browser will open. Select the data files you wish to process. You will be prompted if you want to add more files to the queue.
4. **Configure Settings:**
   * **Number of Segments:** The total number of segments in your MS run.
   * **Time per Segment (min):** The duration of each segment. The script provides a detected maximum runtime to help guide this.
   * **Precursor m/z:** The m/z of your parent ion. If left blank, the script will automatically detect the most abundant peak in the first run.
   * **Number of Expected Fragments:** The strict maximum number of fragment columns you want in your final Excel file.
   * **Expected m/z fragments:** A comma-separated list of specific fragment m/z values you want to track. If left blank, the script will automatically identify and track the fragments with the highest total intensity across the run.
5. **Save:** Choose a destination and filename for your output Excel file.

The script will process each file, extract the relevant m/z intensities per segment, and compile everything into a single, multi-run Excel workbook with bold headers for easy reading.
