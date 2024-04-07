import fitz  # PyMuPDF
import pandas as pd
import numpy as np

class PDF:
    def __init__(self, file):
        self.pdf_doc = fitz.open(file)
        
    @staticmethod
    def merge_bins(arr):
        result = []
        n = len(arr)
        i = 0
        while i < n:
            if i + 1 < n and arr[i + 1] - arr[i] < 4:
                result.append((arr[i + 1] + arr[i]) / 2)
                i += 2
            else:
                result.append(arr[i])
                i += 1
        return np.array(result)

    def multipage_table(self, pages, num_cols, skip_rows=0, skip_cols=0, skip_rects=0):
        all_data = []
        for p in pages:
            pdf_page = self.pdf_doc[int(p)]

            x1_values = []
            y1_values = []
            for rect in pdf_page.get_drawings()[skip_rects:]:
                x1_values.append(rect['rect'][0])
                y1_values.append(rect['rect'][1])
                
            x_bins = np.sort(list(set(x1_values)))
            y_bins = np.sort(list(set(y1_values)))
            y_grid = self.merge_bins(y_bins)
            x_grid = self.merge_bins(x_bins)

            for row in range(skip_rows, len(y_grid)-1):
                row_data = []
                y1, y2 = y_grid[row], y_grid[row+1]

                for col in range(skip_cols, num_cols):
                    x1, x2 = x_grid[col], x_grid[col+1]
                    # Extract text within cell boundaries
                    cell_text = pdf_page.get_text("text", clip=(x1, y1, x2, y2))
                    row_data.append(cell_text.replace("\n", " ").strip())
                all_data.append(row_data)
                
        # Create DataFrame without predefined headers
        df = pd.DataFrame(all_data)
        # Optionally, you can set the first row as headers or create generic column names
        # df.columns = df.iloc[0]  # Uncomment if first row should be headers
        # df = df[1:]  # Uncomment if first row should be headers
        return df

# Usage
file_path = r"C:\Users\cmccall\Downloads\Pipe Dimensions and Weight.pdf"
pdf = PDF(file_path)
pages = range(len(pdf.pdf_doc))  # Process all pages or specify the range you need
num_cols = 22  # Adjust based on the number of columns in the table

df = pdf.multipage_table(pages, num_cols, skip_rows=1)  # Adjust skip_rows as needed

# Save or further process your DataFrame `df`
# Example: Save to CSV
df.to_csv('output_data.csv', index=False)
