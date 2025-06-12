import os
import glob
import subprocess
import pandas as pd
import logging
import warnings
warnings.simplefilter("ignore", category=UserWarning)

def latex_escape(text):
    """
    Escapes LaTeX special characters in a string.
    """
    if pd.isna(text):
        return "-"
    return (str(text)
            .replace('_', '\\_')
            .replace('%', '\\%')
            .replace('&', '\\&'))

def build_main_table(test_type, df):
    """
    Returns a LaTeX string for the 'Main Table of Steps'.
    """
    latex_str = rf"""
\begin{{longtable}}{{|p{{3cm}}|p{{10cm}}|p{{3cm}}|}}
\caption{{{test_type} Test-Steps Table}}
\label{{tab:{test_type.replace(' ','_')}_TestSteps}} \\
\rowcolor{{headerblue}}
\textcolor{{white}}{{\textbf{{Test Step ID}}}} & 
\textcolor{{white}}{{\textbf{{Test Step Name}}}} & 
\textcolor{{white}}{{\textbf{{Start Time}}}} \\
\hline
\endfirsthead

\rowcolor{{headerblue}}
\textcolor{{white}}{{\textbf{{Test Step ID}}}} & 
\textcolor{{white}}{{\textbf{{Test Step Name}}}} & 
\textcolor{{white}}{{\textbf{{Start Time}}}} \\
\hline
\endhead
"""
    row_lines = []
    for _, row_data in df.iterrows():
        step_id   = latex_escape(row_data.get("Test Step ID", ""))
        step_name = latex_escape(row_data.get("Test Step Name", ""))
        start_t   = latex_escape(row_data.get("Start/Step Time (NEM)", ""))
        row_lines.append(f"{step_id} & {step_name} & {start_t} \\\\ \\hline")
    latex_str += "\n".join(row_lines)
    latex_str += "\n\\end{longtable}\n"
    return latex_str

def build_comfail_summary(df):
    """
    Returns a custom LaTeX summary table for COMFAIL.
    """
    columns = [
        ("Test Step ID", 3.0),
        ("Communication Fail", 5.0),
        ("Date and Time (NEM)", 4.0),
        ("POC Pref", 2.5),
        ("Wait Time", 2.5),
        ("Ramp Rate", 2.5),
        ("Result (1)", 2.5),
        ("Result (2)", 2.5),
        ("Communication Reinstatement", 4.0),
        ("Inverter Ramp Down, Disconnect, Lockouts", 5.5),
        ("If Switch Offsite", 3.5),
        ("Plot Figure", 3.0),
        ("POC INV1", 3.0),
    ]
    column_spec = "|".join([f"p{{{width}cm}}" for _, width in columns])
    column_spec = f"|{column_spec}|"
    header_cells = [
        rf"\textcolor{{white}}{{\textbf{{{col_name}}}}}"
        for col_name, _ in columns
    ]
    header_row = " & ".join(header_cells) + r" \\ \hline"
    table_rows = []
    for _, row_data in df.iterrows():
        test_step_id = latex_escape(row_data.get("Test Step ID", ""))
        comm_fail    = latex_escape(row_data.get("Test Step Name", ""))
        date_time    = latex_escape(row_data.get("Start/Step Time (NEM)", ""))
        num_blank = len(columns) - 3
        blanks = " &" * num_blank
        row_str = f"{test_step_id} & {comm_fail} & {date_time}" + blanks + r" \\ \hline"
        table_rows.append(row_str)
    latex_str = rf"""
\noindent
\renewcommand{{\arraystretch}}{{1.3}}
\setlength{{\tabcolsep}}{{6pt}}
\begin{{table}}[H]
    \centering
    \caption{{COMFAIL Results Table (Custom, from df)}}
    \label{{tab:COMFAIL_Results}}
    \begin{{tabularx}}{{\linewidth}}{{{column_spec}}}
    \hline
    \rowcolor{{headerblue}}
    {header_row}
"""
    latex_str += "\n".join(table_rows)
    latex_str += r"""
    \end{tabularx}
\end{table}
"""
    return latex_str

def build_default_summary_table(test_type, summary_df):
    """
    Returns the default summary table from 'summary.xlsx'.
    """
    num_cols = len(summary_df.columns)
    if num_cols > 1:
        column_format = "|p{5cm}|" + "X|"*(num_cols-1)
    else:
        column_format = "|p{5cm}|"
    headers_list = [
        rf"\textcolor{{white}}{{\textbf{{{col.replace('_', ' ')}}}}}"
        for col in summary_df.columns
    ]
    header_row = " & ".join(headers_list) + " \\\\ \\hline\n"
    row_lines = []
    for _, row_data in summary_df.iterrows():
        row_values = [latex_escape(val) for val in row_data]
        row_lines.append(" & ".join(row_values) + " \\\\ \\hline")
    latex_str = rf"""
\noindent
\renewcommand{{\arraystretch}}{{1.3}}
\setlength{{\tabcolsep}}{{6pt}}

\begin{{table}}[H]
    \centering
    \caption{{{test_type.replace('_','\\_')} Results Table}}
    \label{{tab:{test_type.replace(' ','_')}_Results}}
    \begin{{tabularx}}{{\linewidth}}{{{column_format}}}
        \hline
        \rowcolor{{headerblue}}
        {header_row}
{ "\n".join(row_lines) }
    \end{{tabularx}}
\end{{table}}
"""
    return latex_str

def write_section_file(file_path, content):
    """
    Writes content to a file.
    """
    with open(file_path, "w", encoding="utf-8") as f:
        f.write(content)

def write_standalone_summary_file(summary_tex_path, summary_code):
    """
    Writes a complete standalone LaTeX file for the summary table.
    This file will use the article class with A3 paper in landscape mode and a 2in margin.
    """
    summary_content = r"""\documentclass[a3paper,landscape]{article}
\usepackage[margin=2in]{geometry}
\usepackage{graphicx}
\usepackage{longtable}
\usepackage{tabularx}
\usepackage[table]{xcolor}
\usepackage{float} % Provides the [H] float option.
\usepackage{pdfpages}
\definecolor{headerblue}{HTML}{002060}
\pagestyle{empty}
\begin{document}
""" + summary_code + r"""
\end{document}
"""
    with open(summary_tex_path, "w", encoding="utf-8") as f:
        f.write(summary_content)


def compile_latex(tex_file_path):
    try:
        subprocess.run(
            ["tectonic", tex_file_path, "--synctex", "--keep-logs"],
            cwd=os.path.dirname(tex_file_path),
            check=True,
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE
        )
        return True
    except subprocess.CalledProcessError as e:
        print("Error compiling", tex_file_path)
        print(e.stdout.decode())
        print(e.stderr.decode())
        return False

def generate_main_tex(main_tex_path, project_name, section_files, exec_summary_file):
    """
    Generates main.tex that includes all section files.
    The main.tex file is saved outside the Matter folder, and the section files are referenced via their relative path.
    """
    # Relative path for the executive summary file inside the Matter folder (using forward slashes)
    exec_summary_relative = os.path.join("Matter", os.path.basename(exec_summary_file))
    exec_summary_relative = exec_summary_relative.replace("\\", "/")
    
    main_content = rf"""\documentclass{{article}}
\usepackage{{graphicx}}
\usepackage{{longtable}}
\usepackage{{underscore}}
\usepackage{{float}}
\usepackage[a4paper, margin=1in]{{geometry}}
\usepackage{{tocbibind}}
\usepackage{{hyperref}}
\usepackage{{pdflscape}}
\usepackage{{tabularx}}
\usepackage{{fancyhdr}}
\usepackage[table]{{xcolor}}
\usepackage{{pdfpages}}

\definecolor{{headerblue}}{{HTML}}{{002060}}

\title{{Hold Point Testing Report}}
\author{{{project_name.replace('_', '\\_')}}}
\date{{\today}}

\pagestyle{{fancy}}
\fancyhf{{}}
\fancyhead[L]{{HP2 Test Report | {project_name.replace('_', '\\_')}}}
\fancyhead[R]{{\includegraphics[height=0.8cm]{{Figures/Theme/company_logo.jpg}}}}
\renewcommand{{\headrulewidth}}{{0.2pt}}
\setlength{{\headsep}}{{35pt}}
\fancyfoot[C]{{\thepage}}
\renewcommand{{\footrulewidth}}{{0pt}}

\hypersetup{{
    colorlinks=true,
    linkcolor=blue,
    filecolor=magenta,
    urlcolor=cyan,
    pdftitle={{{{Hold Point Testing Report}}}},
    bookmarks=true,
    pdfpagemode=FullScreen,
}}

\begin{{document}}

\maketitle
\tableofcontents
\newpage
\listoftables
\newpage
\listoffigures

\newpage
\section*{{Executive Summary}}
\input{{{exec_summary_relative}}}

\newpage
"""
    for section in section_files:
        main_content += rf"\include{{{section}}}" + "\n"
    main_content += "\n\\end{{document}}\n"
    
    with open(main_tex_path, "w", encoding="utf-8") as f:
        f.write(main_content)

def generate_latex_report(
    excel_path, 
    project_name, 
    output_tex_path, 
    overlays_dir="Overlays",
    plots_dir="Plots"
):

    try:
        # Determine the base directory and create the Matter folder
        base_dir = os.path.dirname(output_tex_path)
        matter_dir = os.path.join(base_dir, "Matter")
        os.makedirs(matter_dir, exist_ok=True)
        
        log_file_path = os.path.join(matter_dir, "report_generation.log")
        logging.basicConfig(
            filename=log_file_path,
            level=logging.INFO,
            format="%(asctime)s - %(levelname)s - %(message)s"
        )
        
        # Load the Excel file and group tests by "Test Type"
        xls = pd.ExcelFile(excel_path)
        tests_df = xls.parse('Tests')
        test_types = tests_df["Test Type"].dropna().unique()
        test_sections = {test_type: tests_df[tests_df["Test Type"] == test_type]
                         for test_type in test_types}
        
        section_files = []
        
        # Write Executive Summary to its own file in the Matter folder
        exec_summary_content = rf"""This report documents the Hold Point Testing for {project_name.replace('_', '\\_')}.
Below is the breakdown of test steps categorized by test type."""
        exec_summary_file = os.path.join(matter_dir, "section_executive_summary.tex")
        write_section_file(exec_summary_file, exec_summary_content)
        
        # Process each test type into its own subfolder with its section file (and standalone summary if available)
        for test_type, df in test_sections.items():
            # Create a file-safe version of the test type name
            file_safe_test_type = test_type.replace(" ", "_").replace("/", "_")
            # Create subfolder for this test type inside Matter
            test_folder = os.path.join(matter_dir, file_safe_test_type)
            os.makedirs(test_folder, exist_ok=True)
            
            section_filename = f"section_{file_safe_test_type}.tex"
            section_filepath = os.path.join(test_folder, section_filename)
            # For \include, store the relative path from the main file (without .tex extension)
            relative_section_path = os.path.join("Matter", file_safe_test_type, section_filename[:-4])
            relative_section_path = relative_section_path.replace("\\", "/")
            section_files.append(relative_section_path)
            
            section_content = ""
            section_content += f"\\section{{{test_type.replace('_', '\\_')}}}\n"
            section_content += build_main_table(test_type, df)
            
            # Check for summary table: either COMFAIL (custom) or default summary if summary.xlsx exists
            summary_excel_path = os.path.join(
                r"C:\Users\rodrigo.ghersi\OneDrive - EPEC Group Pty Ltd\Jupyter Notebooks\00. Projects\KSF_Processed",
                "SF" + test_type,
                "POC",
                "summary.xlsx"
            )
            
            # If there's a summary table, create a standalone summary document
            if test_type == "COMFAIL" or os.path.exists(summary_excel_path):
                # Create standalone summary file (summary_standalone.tex) in the same folder
                summary_tex_path = os.path.join(test_folder, "summary_standalone.tex")
                
                if test_type == "COMFAIL":
                    summary_code = build_comfail_summary(df)
                else:
                    summary_xls = pd.ExcelFile(summary_excel_path)
                    summary_df = summary_xls.parse('Summary').fillna("-")
                    summary_code = build_default_summary_table(test_type, summary_df)
                
                write_standalone_summary_file(summary_tex_path, summary_code)
                # Compile the standalone summary document to produce a PDF
                compile_success = compile_latex(summary_tex_path)
                if not compile_success:
                    print(f"Failed to compile summary for test type {test_type}")
                
                # In the section file, include the compiled summary PDF using pdfpages.
                section_content += r"\subsection{Results}" + "\n"
                section_content += rf"""
\clearpage
\includepdf[pages=-, pagecommand={\thispagestyle{{empty}}}, fitpaper=true]{{Matter/{file_safe_test_type}/summary_standalone.pdf}}
\clearpage
"""
            # Overlays Subsection
            if "Overlay" in df.columns:
                overlays_df = df[df["Overlay"].astype(str).str.strip().str.lower() != "no"]
                if not overlays_df.empty:
                    section_content += r"\subsection{Overlays}" + "\n"
                    for _, row in overlays_df.iterrows():
                        step_id_escaped = latex_escape(row.get('Test Step ID', ""))
                        image_path = os.path.join(overlays_dir, f"{row['Test Step ID']}.png")
                        image_path = image_path.replace("\\", "/")
                        section_content += rf"""
\begin{{figure}}[H]
    \centering
    \includegraphics[width=\textwidth]{{{image_path}}}
    \caption{{Test Step {step_id_escaped} Overlay}}
    \label{{fig:{row['Test Step ID'].replace('_', '').replace('%', '')}}}
\end{{figure}}
"""
            # Plots Subsection
            pattern = os.path.join(plots_dir, f"*{test_type}*.*")
            matched_files = glob.glob(pattern)
            if matched_files:
                section_content += r"\subsection{Plots}" + "\n"
                for plot_file in matched_files:
                    base_name = os.path.basename(plot_file)
                    latex_path = os.path.join("Figures/Plots", base_name).replace("\\", "/")
                    section_content += rf"""
\begin{{figure}}[H]
    \centering
    \includegraphics[width=\textwidth]{{{latex_path}}}
    \caption{{Test Type {test_type} Plot: {base_name}}}
    \label{{fig:plot_{test_type}_{base_name.replace('.', '_').replace(' ', '_')}}}
\end{{figure}}
"""
            write_section_file(section_filepath, section_content)
        
        # Generate main.tex outside the Matter folder (at the output_tex_path)
        generate_main_tex(output_tex_path, project_name, section_files, exec_summary_file)
        
        print(f"LaTeX report generated successfully: {output_tex_path}")
        logging.info(f"LaTeX report generated successfully: {output_tex_path}")
    
    except Exception as e:
        logging.error(f"Error generating LaTeX report: {e}")
        print(f"Error: {e}")

# Example usage
if __name__ == "__main__":
    excel_file_path = r"C:\Users\rodrigo.ghersi\OneDrive - EPEC Group Pty Ltd\Jupyter Notebooks\04. Report Generation\Projects\KSF\Kingaroy SF HP2 Testing.xlsx"
    project_name = "Kingaroy SF"
    # Main .tex file path is outside the Matter folder
    output_tex_file = r"C:\Users\rodrigo.ghersi\OneDrive - EPEC Group Pty Ltd\Jupyter Notebooks\04. Report Generation\Projects\KSF\Kingaroy_SF_HP2_Testing.tex"
    
    generate_latex_report(
        excel_file_path, 
        project_name, 
        output_tex_file,
        overlays_dir="Figures/Overlays",
        plots_dir=r"C:\Users\rodrigo.ghersi\OneDrive - EPEC Group Pty Ltd\Jupyter Notebooks\04. Report Generation\Projects\KSF\Figures\Plots"
    )
