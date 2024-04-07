from openpyxl import load_workbook


excel_file = "/content/analysis.xlsx"  
wb = load_workbook(excel_file)
ws_data_sheet = wb["DataSheet"]
ws_coding_sheet = wb["CodingSheet"]

# a. Calculate the "Student Score"
def calculate_student_score(data_sheet, coding_sheet):
    student_scores = []
    for index, data_row in data_sheet.iterrows():
        coding_row = coding_sheet.iloc[index]
        score = [1 if data_resp == correct_resp else 0 for data_resp, correct_resp in zip(data_row, coding_row)]
        student_scores.append(score)
    return student_scores

# b. Calculate the test score of each student
def calculate_test_score(student_scores):
    test_scores = [sum(score) for score in student_scores]
    return test_scores

# c. Calculate the percentage correct for each student
def calculate_percentage_correct(student_scores):
    num_questions = len(student_scores[0])
    percentage_correct = [(sum(score) / num_questions) * 100 for score in student_scores]
    return percentage_correct

# d. Calculate the level of each student based on percentage correct
def calculate_student_level(percentage_correct):
    student_levels = []
    for percentage in percentage_correct:
        if percentage >= 90:
            level = "Excellent"
        elif percentage >= 70:
            level = "Good"
        elif percentage >= 50:
            level = "Average"
        else:
            level = "Poor"
        student_levels.append(level)
    return student_levels

# Perform calculations
student_scores = calculate_student_score(data_sheet, coding_sheet)
test_scores = calculate_test_score(student_scores)
percentage_correct = calculate_percentage_correct(student_scores)
student_levels = calculate_student_level(percentage_correct)

# Output results
results_df = pd.DataFrame({
    'Test Score': test_scores,
    'Percentage Correct': percentage_correct,
    'Level': student_levels
})
print(results_df)
