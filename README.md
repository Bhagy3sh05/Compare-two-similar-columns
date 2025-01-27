Fuzzy Matcher: Intelligent String Matching Tool

>>Overview
This Python-based tool enables intelligent matching between similar but non-identical text entries across datasets. It functions like an enhanced VLOOKUP that can identify and match near-identical entries. The project was developed with assistance from Cursor AI.
Repository Contents

Fuzzy.py: Initial implementation for specific use cases
fuzzy_matcher_hybrid.py: Enhanced version with GUI, capable of comparing Excel columns
Executable file ( FuzzyMatcher.exe): Ready-to-use application for non-technical users ( dropbox link: https://www.dropbox.com/scl/fi/m6pezhylyqzch430ktqqx/FuzzyMatcher.exe?rlkey=dvcs1ol1unh2ikd6905efrl0s&st=uk3nxx6g&dl=0)

>> Usage Instructions

Launch FuzzyMatcher.exe

Click "Browse" to select your Excel file containing the columns for comparison
Note: If your data is in separate files, combine the relevant columns into a single Excel file


Enter the sheet name containing your data

Select the column names from the dropdown menus
Unnamed columns will appear in order. 

Click "Process Fuzzy Matching"

Choose the save location for your output file when prompted

>>Performance Considerations

Match accuracy depends on the similarity between your datasets
Processing time increases with data volume
For large datasets, consider processing in smaller batches
Match quality correlates with the logical consistency between datasets

>>Technical Notes
The variable naming convention in fuzzy_matcher_hybrid.py maintains compatibility with the original Fuzzy.py implementation. While this may affect code readability, it ensures functional consistency across versions.
