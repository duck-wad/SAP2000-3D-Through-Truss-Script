# SAP2000-3D-Through-Truss-Script

Requirements: tqdm, comtypes, pandas, numpy, matplotlib

1. Put the BASE.sdb model in the same folder as main.py
2. Run main.py
3. It will create an output.xlsx with the results of each model iteration. This excel file updates every 10 iterations, so do not open it because it may need access. If you need to check the results midway into the run, make a copy of the excel file
4. Once main.py is finished running, run interpret_results.py which reads from the output.xlsx. This will create a plots folder with the required plots
