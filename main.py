import os
import openai
import pandas as pd

# Setup OpenAI
openai.api_key = os.getenv('OPENAI_API_KEY')
client = openai.OpenAI(
    api_key=openai.api_key,
    base_url="http://localhost:11434/v1"
)

# Define the path to your Excel file
excel_file_path = '/path/to/your/downloads/folder/youtube_comments.xlsx'  # Update this path

# Load the Excel file
df = pd.read_excel(excel_file_path)

def classify_comment(comment):
    try:
        response = client.chat.completions.create(
            model="llama2",
            messages=[
                {"role": "system", "content": """
                    You are a helpful assistant designed to classify comments into: 
                    Neutral, Request, Question, Appreciation, Error or Other. 
                    Must respond with one word.
                    """
                },
                {"role": "user", "content": "This video was really helpful, thanks!"},
                {"role": "assistant", "content": "Appreciation"},
                {"role": "user", "content": "Can you explain how quantum computing works?"},
                {"role": "assistant", "content": "Question"},
                {"role": "user", "content": "Please create a video on how to use the new software."},
                {"role": "assistant", "content": "Request"},
                {"role": "user", "content": "There is an error in the code"},
                {"role": "assistant", "content": "Error"},
                {"role": "user", "content": "This video is okay, nothing special."},
                {"role": "assistant", "content": "Neutral"},
                {"role": "user", "content": comment}  # This is the comment you want to classify
            ]
        )
        category = response.choices[0].message.content.strip()
        print(f"Comment: '{comment}' -> Category: '{category}'")
        return category.title()
    except Exception as e:
        print(f"Error in classifying comment: {e}")
        return "Process Failed"

# Iterate over each row in the DataFrame
for index, row in df.iterrows():
    comment = row['Comment']  # Assuming 'Comment' is the column name for comments
    category = classify_comment(comment)
    # Update the DataFrame with the classification
    df.at[index, 'Category'] = category  # 'Category' is the new column for classifications

# Save the updated DataFrame back to an Excel file
df.to_excel(excel_file_path, index=False)  # Set `index=False` to avoid adding an unnamed index column in the Excel file

print("The Excel file has been updated with categories for each comment.")
