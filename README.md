# SUPERHANS

**S**preadsheet **U**tility for **P**recision **E**xtraction and **R**eporting of **H**igh-**A**ccuracy **N**umerical **S**lopes

## Usage:

Put a placeholder formula called `DERIVATIVE` in an arbitrary unused cell, for each derivative to appear. Don't worry if you get a `#NAME?` error. For example, if you want the derivative of A1 with respect to B1 to appear in cell C1, enter `=DERIVATIVE(C1,A1,B1)` into any cell. Then, run:

```bash
python3 superhans.py input_file.xlsx output_file.xlsx
```