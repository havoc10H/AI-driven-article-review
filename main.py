import tkinter as tk
from tkinter import filedialog, scrolledtext, messagebox
from docx import Document
import pandas as pd
from openai import OpenAI

client = OpenAI(
)

def review_article(article_text, selected_guidelines, selected_exists):
    results = []
    
    for i, guideline in enumerate(selected_guidelines):
        exist = selected_exists[i].strip().lower()  # Get the corresponding 'exist' value and normalize it
        print(guideline)
        print(exist)
        
        try:
            # Define the expectation message based on the exist value
            if exist == 'yes':
                expectation = f"The article should have {guideline.lower()}."
            elif exist == 'no':
                expectation = f"The article should not have {guideline.lower()}."
            elif exist == 'no relevant':
                expectation = f"The relevance of {guideline.lower()} does not apply to this article."
            else:
                expectation = "No specific expectation defined."

            # Construct the prompt based on the guideline and expectation
            prompt = (f"Analyze this article based on the following guideline:\n\n"
                      f"Guideline: {guideline}.\n"
                      f"Expectation: {expectation}\n\n"
                      f"Article:\n{article_text}\n\nAnalysis:")

            # Call the OpenAI API
            response = client.chat.completions.create(
                model="gpt-3.5-turbo",
                messages=[
                    {"role": "user", "content": prompt}
                ]
            )

            # Extract the analysis from the response
            analysis = response['choices'][0]['message']['content'].strip()

            # Determine compliance based on the analysis and the 'exist' value
            if exist == 'yes':
                compliance = "Yes" if "yes" in analysis.lower() else "No"
            elif exist == 'no':
                compliance = "No" if "yes" in analysis.lower() else "Yes"
            elif exist == 'no relevant':
                compliance = "N/A"  # Not applicable
            else:
                compliance = "Unknown"  # Catch-all for unexpected cases

            # Determine color based on compliance
            if compliance == "Yes":
                color = "green"
            elif compliance == "No":
                color = "red"
            else:
                color = "gray"

            # Append the guideline, analysis, compliance, and color to results
            results.append((guideline, analysis, compliance, color))
        
        except Exception as e:
            print(f"An error occurred: {e}")
            results.append((guideline, "Error occurred", "No", "red"))
    
    return results


def start_review():
    article_text = article_input.get("1.0", tk.END)
    selected_guidelines = [guideline for idx, guideline in enumerate(guidelines) if check_vars[idx].get()]
    selected_exists = [exist for idx, exist in enumerate(exists) if check_vars[idx].get()]
    if not article_text.strip() or not selected_guidelines:
        messagebox.showwarning("Input Error", "Please make sure both the article and at least one guideline are provided.")
        return
    results = review_article(article_text, selected_guidelines, selected_exists)
    result_text.config(state=tk.NORMAL)
    result_text.delete("1.0", tk.END)
    for guideline, result, answer, color in results:
        result_text.insert(tk.END, f"Guideline: {guideline}\nResult: {result}\nAnswer: {answer}\n", ("color_" + color))
    result_text.config(state=tk.DISABLED)
    with open("results.txt", "w") as file:
        file.write("\n".join([f"Guideline: {g}\nResult: {r}\nAnswer: {a}\n" for g, r, a, c in results]))

def load_article():
    file_path = filedialog.askopenfilename(filetypes=[("Word Documents", "*.txt;*.doc;*.docx")])
    if file_path:
        doc = Document(file_path)
        article_text = "\n".join([para.text for para in doc.paragraphs])
        article_input.delete("1.0", tk.END)
        article_input.insert(tk.END, article_text)

def load_guidelines():
    file_path = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xls;*.xlsx")])
    if file_path:
        try:
            # Read the Excel file with specific columns
            df = pd.read_excel(file_path, usecols=['title', 'exist'])
            
            # Clear previous checkbuttons
            for widget in guidelines_canvas_frame.winfo_children():
                widget.destroy()

            # Initialize global variables
            global check_vars
            global guidelines
            global exists

            check_vars = []
            guidelines = []
            exists = []

            # Process DataFrame rows
            for _, row in df.iterrows():
                title = str(row.get('title', '')).strip()
                exist = str(row.get('exist', '')).strip().lower()
                if title and exist.lower() != 'nan' and exist != '':
                    guidelines.append(title)
                    exists.append(exist)

                    var = tk.IntVar(value=1)
                    check_vars.append(var)
                    tk.Checkbutton(guidelines_canvas_frame, text=title+' (' + exist + ')', variable=var).pack(anchor='w')

            # Update scrollbar and canvas
            guidelines_canvas.update_idletasks()
            guidelines_canvas.config(scrollregion=guidelines_canvas.bbox("all"))

        except Exception as e:
            print(f"An error occurred while loading guidelines: {e}")

app = tk.Tk()
app.title("Article Reviewing App")

# Layout
frame = tk.Frame(app)
frame.pack(fill=tk.BOTH, expand=True)

# Left section (Article Input)
left_frame = tk.Frame(frame)
left_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=10, pady=10)

tk.Label(left_frame, text="Article Section").pack()
article_input = scrolledtext.ScrolledText(left_frame, wrap=tk.WORD)
article_input.pack(fill=tk.BOTH, expand=True)

tk.Button(left_frame, text="Load Article", command=load_article).pack()

# Right section (Guidelines with Scrolling)
right_frame = tk.Frame(frame)
right_frame.pack(side=tk.RIGHT, fill=tk.BOTH, expand=True, padx=10, pady=10)

tk.Label(right_frame, text="Guidelines Section").pack()

# Canvas for guidelines
guidelines_canvas = tk.Canvas(right_frame)
guidelines_canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

# Scrollbars for canvas
v_scrollbar = tk.Scrollbar(right_frame, orient=tk.VERTICAL, command=guidelines_canvas.yview)
v_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

h_scrollbar = tk.Scrollbar(right_frame, orient=tk.HORIZONTAL, command=guidelines_canvas.xview)
h_scrollbar.pack(side=tk.BOTTOM, fill=tk.X)

# Frame inside canvas
guidelines_canvas_frame = tk.Frame(guidelines_canvas)
guidelines_canvas.create_window((0, 0), window=guidelines_canvas_frame, anchor="nw")

# Configure canvas scrollbars
guidelines_canvas.configure(yscrollcommand=v_scrollbar.set, xscrollcommand=h_scrollbar.set)

# Button to load guidelines
tk.Button(right_frame, text="Load Guidelines", command=load_guidelines).pack()

# Result section
result_frame = tk.Frame(app)
result_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

result_header_frame = tk.Frame(result_frame)
result_header_frame.pack(fill=tk.X)

tk.Label(result_header_frame, text="Results").pack(side=tk.LEFT)

tk.Button(result_header_frame, text="Start Review", command=start_review).pack(side=tk.RIGHT)

result_text = scrolledtext.ScrolledText(result_frame, wrap=tk.WORD, state=tk.DISABLED)
result_text.pack(fill=tk.BOTH, expand=True, pady=5)

# Add tag configuration for colored text
result_text.tag_configure("color_green", foreground="green")
result_text.tag_configure("color_red", foreground="red")

app.mainloop()
