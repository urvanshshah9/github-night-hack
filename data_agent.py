import os
import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns
from dotenv import load_dotenv
import yagmail

# Load environment variables
load_dotenv()
EMAIL_USER = os.getenv('EMAIL_USER')
EMAIL_PASS = os.getenv('EMAIL_PASS')
EMAIL_RECEIVER = os.getenv('EMAIL_RECEIVER')

def analyze_and_visualize(file_path):
    df = pd.read_excel(file_path)
    summary = df.describe(include='all')
    summary_file = 'summary.txt'
    with open(summary_file, 'w') as f:
        f.write(summary.to_string())

    # Example: plot first two numeric columns if available
    numeric_cols = df.select_dtypes(include='number').columns
    plots = []
    if len(numeric_cols) >= 2:
        plt.figure(figsize=(8,6))
        sns.lineplot(x=numeric_cols[0], y=numeric_cols[1], data=df, marker='o')
        plt.title(f'{numeric_cols[1]} vs {numeric_cols[0]}')
        plt.xlabel(numeric_cols[0])
        plt.ylabel(numeric_cols[1])
        plot_file = f'{numeric_cols[1]}_vs_{numeric_cols[0]}.png'
        plt.savefig(plot_file)
        plt.close()
        plots.append(plot_file)
    elif len(numeric_cols) == 1:
        plt.figure(figsize=(8,6))
        sns.histplot(df[numeric_cols[0]].dropna())
        plt.title(f'Distribution of {numeric_cols[0]}')
        plt.xlabel(numeric_cols[0])
        plt.ylabel('Frequency')
        plot_file = f'distribution_{numeric_cols[0]}.png'
        plt.savefig(plot_file)
        plt.close()
        plots.append(plot_file)

    attachments = [summary_file] + plots
    return attachments

def send_email(attachments):
    yag = yagmail.SMTP(EMAIL_USER, EMAIL_PASS)
    subject = "Automated Chemical Data Analysis Report"
    contents = [
        "Attached are the summary statistics and visualizations from your chemical experiment data."
    ]
    yag.send(EMAIL_RECEIVER, subject, contents, attachments=attachments)
    print("Email sent!")

if __name__ == "__main__":
    file_path = input("Enter the path to your Excel file: ")
    if not os.path.exists(file_path):
        print(f"File not found: {file_path}")
        exit(1)
    attachments = analyze_and_visualize(file_path)
    send_email(attachments)
