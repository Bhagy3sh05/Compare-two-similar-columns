import pandas as pd
from fuzzywuzzy import fuzz
import numpy as np

# Read the Excel file
excel_path = r'C:\Users\hp5cg\OneDrive\Desktop\Fuzzymatch\Ptyhon\123.xlsx'
df = pd.read_excel(excel_path, sheet_name='Sheet2')

# Get lists of stores and customers
stores = df['Unnamed: 1'].dropna().tolist()
customers = df['Unnamed: 2'].dropna().tolist()

def calculate_similarity(str1, str2):
    str1, str2 = str(str1).lower(), str(str2).lower()
    ratio = fuzz.ratio(str1, str2)
    partial_ratio = fuzz.partial_ratio(str1, str2)
    token_sort_ratio = fuzz.token_sort_ratio(str1, str2)
    token_set_ratio = fuzz.token_set_ratio(str1, str2)
    return max(ratio, partial_ratio, token_sort_ratio, token_set_ratio)

# Create a similarity matrix
similarity_matrix = []
for store in stores:
    row = []
    for customer in customers:
        score = calculate_similarity(store, customer)
        row.append(score)
    similarity_matrix.append(row)

# Convert to numpy array for easier manipulation
similarity_matrix = np.array(similarity_matrix)

results = []
used_store_indices = set()
used_customer_indices = set()

# While we can still make matches
while len(results) < min(len(stores), len(customers)):
    # Find the highest remaining score
    max_score = -1
    best_store_idx = -1
    best_customer_idx = -1
    
    for i in range(len(stores)):
        for j in range(len(customers)):
            if (i not in used_store_indices and 
                j not in used_customer_indices and 
                similarity_matrix[i][j] > max_score):
                max_score = similarity_matrix[i][j]
                best_store_idx = i
                best_customer_idx = j
    
    if best_store_idx == -1 or best_customer_idx == -1:
        break
        
    # Add this match to results
    results.append({
        'Store': stores[best_store_idx],
        'Customer': customers[best_customer_idx],
        'Match Score': max_score
    })
    
    # Mark these indices as used
    used_store_indices.add(best_store_idx)
    used_customer_indices.add(best_customer_idx)

# Create DataFrame from results
result_df = pd.DataFrame(results)

# Sort by match score in descending order
result_df = result_df.sort_values('Match Score', ascending=False)

# Add unmatched stores
unmatched_stores = [store for i, store in enumerate(stores) if i not in used_store_indices]
if unmatched_stores:
    unmatched_df = pd.DataFrame({
        'Store': unmatched_stores,
        'Customer': ['NO MATCH FOUND'] * len(unmatched_stores),
        'Match Score': [0] * len(unmatched_stores)
    })
    result_df = pd.concat([result_df, unmatched_df], ignore_index=True)

# Save results to Excel
output_path = r'C:\Users\hp5cg\OneDrive\Desktop\Fuzzymatch\Ptyhon\fuzzy_match_results.xlsx'
result_df.to_excel(output_path, index=False)

# Print summary
print("\nTop 10 matches:")
print(result_df.head(10))
print(f"\nTotal matches found: {len(results)}")
print(f"Number of unmatched stores: {len(unmatched_stores)}")
print("\nResults saved to fuzzy_match_results.xlsx")

# Print perfect matches
perfect_matches = result_df[result_df['Match Score'] == 100]
if not perfect_matches.empty:
    print("\nPerfect Matches (100% score):")
    print(perfect_matches[['Store', 'Customer']])