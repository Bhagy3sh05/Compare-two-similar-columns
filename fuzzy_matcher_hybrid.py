import tkinter as tk
from tkinter import filedialog, ttk, messagebox
import pandas as pd
from fuzzywuzzy import fuzz
import numpy as np
import os
import time
from threading import Thread

class FuzzyMatcherApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Fuzzy Matcher")
        self.root.geometry("600x500")
        
        # Style
        self.style = ttk.Style()
        self.style.configure('TButton', padding=5)
        self.style.configure('TLabel', padding=5)
        
        # Variables
        self.input_file_path = tk.StringVar()
        self.sheet_name = tk.StringVar(value="Sheet1")
        self.column1 = tk.StringVar()
        self.column2 = tk.StringVar()
        
        # Create GUI elements
        self.create_widgets()
        
    def create_widgets(self):
        # Input File Section
        input_frame = ttk.LabelFrame(self.root, text="Input File", padding=10)
        input_frame.pack(fill="x", padx=10, pady=5)
        
        ttk.Label(input_frame, text="Excel File:").pack(anchor="w")
        
        file_frame = ttk.Frame(input_frame)
        file_frame.pack(fill="x")
        
        self.file_entry = ttk.Entry(file_frame, textvariable=self.input_file_path)
        self.file_entry.pack(side="left", fill="x", expand=True, padx=(0, 5))
        
        ttk.Button(file_frame, text="Browse", command=self.browse_file).pack(side="right")
        
        # Sheet Name
        ttk.Label(input_frame, text="Sheet Name:").pack(anchor="w")
        ttk.Entry(input_frame, textvariable=self.sheet_name).pack(fill="x")
        
        # Column Selection Frame
        col_frame = ttk.LabelFrame(self.root, text="Column Selection", padding=10)
        col_frame.pack(fill="x", padx=10, pady=5)
        
        # First Column
        ttk.Label(col_frame, text="First Column Name:").pack(anchor="w")
        self.col1_combo = ttk.Combobox(col_frame, textvariable=self.column1)
        self.col1_combo.pack(fill="x")
        
        # Second Column
        ttk.Label(col_frame, text="Second Column Name:").pack(anchor="w")
        self.col2_combo = ttk.Combobox(col_frame, textvariable=self.column2)
        self.col2_combo.pack(fill="x")
        
        # Process Button
        ttk.Button(self.root, text="Process Fuzzy Matching", 
                  command=self.process_matching).pack(pady=20)
        
        # Progress Frame
        progress_frame = ttk.LabelFrame(self.root, text="Progress", padding=10)
        progress_frame.pack(fill="x", padx=10, pady=5)
        
        # Progress Bar
        self.progress_var = tk.DoubleVar()
        self.progress_bar = ttk.Progressbar(progress_frame, 
                                          variable=self.progress_var,
                                          maximum=100)
        self.progress_bar.pack(fill="x", pady=5)
        
        # Status Labels
        self.status_label = ttk.Label(progress_frame, text="", wraplength=550)
        self.status_label.pack(pady=5)
        
        self.comparison_label = ttk.Label(progress_frame, text="", wraplength=550)
        self.comparison_label.pack(pady=5)
        
    def browse_file(self):
        file_path = filedialog.askopenfilename(
            filetypes=[("Excel files", "*.xlsx *.xls")]
        )
        if file_path:
            self.input_file_path.set(file_path)
            self.update_column_lists()
    
    def update_column_lists(self):
        try:
            df = pd.read_excel(self.input_file_path.get(), sheet_name=self.sheet_name.get())
            columns = df.columns.tolist()
            self.col1_combo['values'] = columns
            self.col2_combo['values'] = columns
        except Exception as e:
            messagebox.showerror("Error", f"Error reading file: {str(e)}")
    
    def update_progress(self, current, total, message="Processing..."):
        progress = (current / total) * 100
        self.progress_var.set(progress)
        self.status_label.config(text=f"{message} ({progress:.1f}% complete)")
        self.comparison_label.config(
            text=f"Processed {current:,} of {total:,} comparisons")
        self.root.update()
    
    def calculate_similarity(self, str1, str2):
        # Handle empty or null values
        if pd.isna(str1) or pd.isna(str2) or str1 == '' or str2 == '':
            return 0
        
        # Convert to lowercase for comparison
        str1, str2 = str(str1).lower(), str(str2).lower()
        
        # Exact match
        if str1 == str2:
            return 100
        
        # Calculate basic ratios
        ratio = fuzz.ratio(str1, str2)
        partial_ratio = fuzz.partial_ratio(str1, str2)
        token_sort_ratio = fuzz.token_sort_ratio(str1, str2)
        token_set_ratio = fuzz.token_set_ratio(str1, str2)
        
        # Get the maximum of all ratios
        max_ratio = max(ratio, token_sort_ratio, token_set_ratio)
        
        # If we have a very high match, return it
        if max_ratio > 95 or partial_ratio == 100:
            return max_ratio
        
        # For other cases, use a weighted approach
        len_diff = abs(len(str1) - len(str2))
        if len_diff > 5:
            # For very different lengths, rely more on partial matching
            return max(partial_ratio * 0.9, max_ratio * 0.8)
        else:
            # For similar lengths, use the higher score
            return max(max_ratio, partial_ratio * 0.95)
    
    def process_matching(self):
        try:
            # Disable the process button
            for child in self.root.winfo_children():
                if isinstance(child, ttk.Button):
                    child.configure(state='disabled')
            
            self.status_label.config(text="Reading input file...")
            self.root.update()
            
            # Read input file
            df = pd.read_excel(self.input_file_path.get(), sheet_name=self.sheet_name.get())
            
            # Get lists of values
            stores = df[self.column1.get()].tolist()
            customers = df[self.column2.get()].tolist()
            
            total_comparisons = len(stores) * len(customers)
            current_comparison = 0
            
            self.status_label.config(text="Calculating similarities...")
            self.root.update()
            
            # Calculate similarity matrix with progress updates
            similarity_matrix = []
            for i, store in enumerate(stores):
                row = []
                for customer in customers:
                    score = self.calculate_similarity(str(store), str(customer))
                    row.append(score)
                    current_comparison += 1
                    if current_comparison % 10 == 0:
                        self.update_progress(current_comparison, total_comparisons)
                similarity_matrix.append(row)
            
            similarity_matrix = np.array(similarity_matrix)
            
            self.status_label.config(text="Finding best matches...")
            self.comparison_label.config(text="")
            self.root.update()
            
            # Find best matches
            results = []
            used_store_indices = set()
            used_customer_indices = set()
            
            # First pass: Find exact and near-exact matches (score >= 95)
            for threshold in [100, 95]:
                for i in range(len(stores)):
                    if i not in used_store_indices:
                        for j in range(len(customers)):
                            if (j not in used_customer_indices and 
                                similarity_matrix[i][j] >= threshold):
                                results.append({
                                    self.column1.get(): stores[i],
                                    self.column2.get(): customers[j],
                                    'Match Score': similarity_matrix[i][j]
                                })
                                used_store_indices.add(i)
                                used_customer_indices.add(j)
                                break
            
            # Second pass: Find remaining best matches
            while len(results) < min(len(stores), len(customers)):
                max_score = -1
                best_store_idx = -1
                best_customer_idx = -1
                
                for i in range(len(stores)):
                    if i not in used_store_indices:
                        for j in range(len(customers)):
                            if (j not in used_customer_indices and 
                                similarity_matrix[i][j] > max_score):
                                max_score = similarity_matrix[i][j]
                                best_store_idx = i
                                best_customer_idx = j
                
                if best_store_idx == -1 or best_customer_idx == -1 or max_score < 20:
                    break
                    
                results.append({
                    self.column1.get(): stores[best_store_idx],
                    self.column2.get(): customers[best_customer_idx],
                    'Match Score': max_score
                })
                
                used_store_indices.add(best_store_idx)
                used_customer_indices.add(best_customer_idx)
                
                self.update_progress(
                    len(results), 
                    min(len(stores), len(customers)), 
                    "Finding best matches..."
                )
            
            # Create results DataFrame
            result_df = pd.DataFrame(results)
            
            # Add original variation column if it exists
            if 'Original Variation' in df.columns:
                result_df = pd.merge(
                    result_df,
                    df[[self.column1.get(), 'Original Variation']],
                    on=self.column1.get(),
                    how='left'
                )
            
            self.status_label.config(text="Saving results...")
            self.comparison_label.config(text="")
            self.root.update()
            
            # Get output file location
            output_path = filedialog.asksaveasfilename(
                defaultextension=".xlsx",
                filetypes=[("Excel files", "*.xlsx")],
                title="Save Results As"
            )
            
            if output_path:
                result_df.to_excel(output_path, index=False)
                self.status_label.config(
                    text=f"Processing complete! Results saved to: {output_path}")
                self.comparison_label.config(
                    text=f"Total matches found: {len(results)}")
                messagebox.showinfo("Success", 
                    f"Matching complete!\nTotal matches found: {len(results)}")
            
        except Exception as e:
            messagebox.showerror("Error", f"An error occurred: {str(e)}")
        finally:
            # Re-enable the process button
            for child in self.root.winfo_children():
                if isinstance(child, ttk.Button):
                    child.configure(state='normal')

if __name__ == "__main__":
    root = tk.Tk()
    app = FuzzyMatcherApp(root)
    root.mainloop()