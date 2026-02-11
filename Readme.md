Security and Reliability Analysis of AI-Generated C++ Code (ECC + CWE)

This repository contains the code and results for the paper:

Security and Reliability Analysis of AI-Generated C++ Code Using Static Analysis: An Empirical Evaluation Based on Saudi NCA Essential Cybersecurity Controls (ECC)

Repository Structure

code/
Scripts used to generate C++ code from LLMs:

chatgpt.py

claude.py

deepseek.py

gemini.py

results/
Exported tables and summary outputs:

tables/ → CSV files for Tables 3–10

summaries/ → Cppcheck and Clang-Tidy summary outputs

figures/
Figures used in the paper (Figure 1–6)

Requirements

Python 3.11+

Cppcheck 2.17.1

Clang-Tidy 21.1.8

CMake 4.2.1 (for generating compilation databases)

How to Run

Run the script for the target model:

python3 code/chatgpt.py
python3 code/claude.py
python3 code/deepseek.py
python3 code/gemini.py


Each script generates C++ programs for the defined tasks and stores outputs for analysis.

Static Analysis

The generated C++ programs were analysed using:

Cppcheck

Clang-Tidy

The final aggregated results are available under results/ and the figures are under figures/.

Citation

If you use this repository, please cite the paper.
