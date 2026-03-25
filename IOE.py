import pandas as pd
import numpy as np

def allocate_students(students_file, programs_file, output_file="allocation_results.txt"):
    # 1. Read the Excel files
    df_students = pd.read_excel(students_file)
    df_programs = pd.read_excel(programs_file)

    # Clean column names: force them to strings to avoid 'float' errors, then strip spaces
    df_students.columns = [str(col).strip() for col in df_students.columns]
    df_programs.columns = [str(col).strip() for col in df_programs.columns]

    # Sort students by Rank to ensure strict priority allocation
    df_students = df_students.sort_values('Rank').reset_index(drop=True)

    # 2. Create a mapping of Program No to Program Name
    program_map = dict(zip(df_programs['Program No'], df_programs['Program Name']))

    # 3. Initialize Quotas for each program
    # Odd numbers: 4 Male, 1 Female
    # Even numbers: 5 Male, 1 Female
    quotas = {}
    for prog_no in df_programs['Program No']:
        if pd.isna(prog_no): # Skip empty rows if any
            continue
        
        prog_no = int(prog_no)
        if prog_no % 2 != 0: # Odd
            quotas[prog_no] = {'Male': 4, 'Female': 1}
        else: # Even
            quotas[prog_no] = {'Male': 5, 'Female': 1}

    # Data structures to store the results
    program_allocations = {name: [] for name in program_map.values()}
    student_allocations = []

    # Find all Priority columns dynamically (P1, P2, P3... P15) safely
    priority_cols = [col for col in df_students.columns if col.startswith('P') and col[1:].isdigit()]

    # 4. Allocation Logic
    for index, row in df_students.iterrows():
        name = str(row['Applicant Name']).strip()
        rank = row['Rank']
        
        # Clean gender string to prevent matching errors
        gender = str(row['Gender']).strip().capitalize() 
        
        assigned_program = "None" # Default if they get nothing

        # Iterate through their priorities
        for p_col in priority_cols:
            prog_no = row[p_col]
            
            # Check if priority is not empty (NaN)
            if pd.notna(prog_no):
                prog_no = int(prog_no)
                
                # Verify program exists in our list
                if prog_no in quotas:
                    # Check if quota is available for the student's gender
                    if quotas[prog_no].get(gender, 0) > 0:
                        
                        # Assign the seat
                        quotas[prog_no][gender] -= 1
                        prog_name = program_map[prog_no]
                        assigned_program = prog_name
                        
                        # Record the assignment for the program list
                        program_allocations[prog_name].append((name, rank))
                        
                        # Stop checking priorities for this student since they got a seat
                        break 
        
        # Record the assignment for the student list
        student_allocations.append({'Name': name, 'Rank': rank, 'Assigned Program': assigned_program})

    # 5. Output Generation (To Console AND Text File)
    with open(output_file, "w", encoding="utf-8") as f:
        
        # Helper function to print to terminal and write to file at the same time
        def log(text=""):
            print(text)
            f.write(str(text) + "\n")

        log("="*60)
        log("OUTPUT 1: ADMITTED STUDENTS PER PROGRAM")
        log("="*60)
        for prog_name, students in program_allocations.items():
            log(f"\nProgram: {prog_name}")
            if not students:
                log("  (No students admitted)")
            else:
                for student_name, student_rank in students:
                    log(f"  - {student_name} (Rank: {student_rank})")

        log("\n" + "="*60)
        log("OUTPUT 2: ASSIGNED PROGRAM PER STUDENT")
        log("="*60)
        
        # Convert to dataframe for clean tabular printing
        results_df = pd.DataFrame(student_allocations)
        log(results_df.to_string(index=False)) 
    
    print(f"\n>>> Success! The output has also been saved to '{output_file}' in your current folder.")

# --- RUN THE CODE ---
if __name__ == "__main__":
    allocate_students("students.xlsx", "programs.xlsx")