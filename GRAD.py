import pandas as pd
import random
from collections import defaultdict

# Load the Excel files
judges_file_path = r"C:\Users\Yash Mehta\Downloads\Example_list_judges.xlsx"
posters_file_path = r"C:\Users\Yash Mehta\Downloads\Sample_input_abstracts.xlsx"
expertise_file_path = r"C:\Users\Yash Mehta\Downloads\book1.csv"

# Read the data into DataFrames
judges_df = pd.read_excel(judges_file_path, sheet_name="Sheet1", engine='openpyxl')
posters_df = pd.read_excel(posters_file_path, sheet_name="Sheet1", engine='openpyxl')
expertise_df = pd.read_csv(expertise_file_path)

def match_expertise_to_poster(poster_keywords, judge_expertise):
    # Match the judge's expertise with the poster's keywords (simple substring check)
    return any(keyword.lower() in judge_expertise.lower() for keyword in poster_keywords.split(','))

def assign_judges_to_posters(judges_df, posters_df, expertise_df):
    # Organize judges by availability
    judges_hour_1 = judges_df[judges_df['Hour available'].isin([1, 'both'])]
    judges_hour_2 = judges_df[judges_df['Hour available'].isin([2, 'both'])]

    # Create a set of advisors to avoid conflicts
    advisor_set = {(row['Advisor FirstName'], row['Advisor LastName']) for _, row in posters_df.iterrows()}

    # Initialize judge assignments
    assignments = defaultdict(list)
    judge_workload = defaultdict(list)  # Stores posters each judge is assigned to

    unassigned_posters = []  # Track posters that couldn't be assigned initially

    for _, poster in posters_df.iterrows():
        poster_num = poster['Poster #']
        advisor = (poster['Advisor FirstName'], poster['Advisor LastName'])
        hour = 1 if poster_num % 2 == 1 else 2
        available_judges = judges_hour_1 if hour == 1 else judges_hour_2

        # Extract the poster keywords or abstract (assuming there's a column named 'Abstract' or 'Keywords')
        poster_keywords = poster['Abstract'] if 'Abstract' in poster else poster['Keywords']
        
        # Filter valid judges based on expertise matching and availability
        valid_judges = []
        expertise_match_found = False

        for j in available_judges['Judge'].tolist():
            judge_info = judges_df.loc[judges_df['Judge'] == j, ['Judge FirstName', 'Judge LastName']].values
            if len(judge_info) > 0:
                judge_tuple = tuple(judge_info[0])
                if len(judge_workload[j]) < 6 and judge_tuple not in advisor_set:
                    # Check if the judge's expertise matches the poster's keywords
                    expertise_match = expertise_df.loc[expertise_df['Judges'] == j, 'Expertise']
                    if not expertise_match.empty:
                        judge_expertise = expertise_match.values[0]
                        if match_expertise_to_poster(poster_keywords, judge_expertise):
                            valid_judges.append(j)
                            expertise_match_found = True

        if not expertise_match_found:
            # If no expertise match is found, try assigning based on similar expertise
            for j in available_judges['Judge'].tolist():
                judge_info = judges_df.loc[judges_df['Judge'] == j, ['Judge FirstName', 'Judge LastName']].values
                if len(judge_info) > 0:
                    judge_tuple = tuple(judge_info[0])
                    if len(judge_workload[j]) < 6 and judge_tuple not in advisor_set:
                        expertise_match = expertise_df.loc[expertise_df['Judges'] == j, 'Expertise']
                        if not expertise_match.empty:
                            judge_expertise = expertise_match.values[0]
                            if poster_keywords.lower() in judge_expertise.lower():  # Similar expertise
                                valid_judges.append(j)

        # If still no valid judges (expertise doesn't match, or no judges with similar expertise)
        if len(valid_judges) < 2:
            # Fallback: Randomly pick judges respecting hour availability
            available_judges = [j for j in available_judges['Judge'].tolist() if len(judge_workload[j]) < 6]
            
            # Check if there are at least 2 available judges
            if len(available_judges) < 2:
                print(f"⚠️ Not enough judges available for Poster {poster_num}, assigning available judges.")
                assigned_judges = available_judges  # Assign all available judges
            else:
                assigned_judges = random.sample(available_judges, 2)

            assignments[poster_num] = assigned_judges
        else:
            # Assign 2 random judges (or from valid ones) while respecting the hour availability
            assigned_judges = random.sample(valid_judges, 2)
            assignments[poster_num] = assigned_judges

        # Update judge workloads
        for judge in assigned_judges:
            judge_workload[judge].append(poster_num)

    return assignments, judge_workload


# Run the assignment process
assignments, judge_workload = assign_judges_to_posters(judges_df, posters_df, expertise_df)

# ============================
# 🔹 Output 1: Posters file with assigned judges
# ============================
posters_df['judge-1'] = posters_df['Poster #'].map(lambda x: assignments.get(x, [None, None])[:2] + [None, None]) \
                                              .apply(lambda lst: lst[0])  # Ensure index 0 exists

posters_df['judge-2'] = posters_df['Poster #'].map(lambda x: assignments.get(x, [None, None])[:2] + [None, None]) \
                                              .apply(lambda lst: lst[1])  # Ensure index 1 exists

posters_output_path = r"C:\Users\Yash Mehta\Downloads\Updated_posters.xlsx"
posters_df.to_excel(posters_output_path, index=False)
print(f"✅ Posters file saved: {posters_output_path}")

# ============================
# 🔹 Output 2: Judges file with assigned posters
# ============================
for i in range(1, 7):  # Create 6 columns for posters
    judges_df[f'poster-{i}'] = judges_df['Judge'].map(lambda j: judge_workload.get(j, [None] * 6)[i - 1] if i <= len(judge_workload.get(j, [])) else None)

judges_output_path = r"C:\Users\Yash Mehta\Downloads\Updated_judges.xlsx"
judges_df.to_excel(judges_output_path, index=False)
print(f"✅ Judges file saved: {judges_output_path}")

# ============================
# 🔹 Output 3: Poster-Judge Assignment Matrix
# ============================
judge_list = judges_df['Judge'].tolist()
poster_list = posters_df['Poster #'].tolist()

# Initialize matrix with zeros
assignment_matrix = pd.DataFrame(0, index=poster_list, columns=judge_list)

# Fill matrix with 1 where a judge is assigned to a poster
for poster, judges in assignments.items():
    for judge in judges:
        assignment_matrix.loc[poster, judge] = 1

matrix_output_path = r"C:\Users\Yash Mehta\Downloads\Assignment_Matrix.xlsx"
assignment_matrix.to_excel(matrix_output_path)
print(f"✅ Assignment matrix saved: {matrix_output_path}")
