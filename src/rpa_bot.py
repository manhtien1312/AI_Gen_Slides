from supabase import create_client
from google import genai
from dotenv import load_dotenv
import pandas as pd
import os

load_dotenv()

# Function to get data from Supabase
def get_data_from_supabase():
    supabase_url = os.getenv("SUPABASE_URL")
    supabase_key = os.getenv("SUPABASE_KEY")
    supabase = create_client(supabase_url, supabase_key)

    response = supabase.schema("public").table('financial_data').select("*").execute()
    df = pd.DataFrame(response.data)
    return df

# Function to get data from local Excel file
def get_data_from_local(file_path):
    df = pd.read_excel(file_path)
    return df

# Function to prepare data for charts
def prepare_data_for_chart(df):
    pie_data = df.groupby('category')['amount'].sum().reset_index()
    pie_data['amount'] = pie_data['amount'].astype(float)

    bar_data = df.groupby(['quarter', 'category'])['amount'].sum().reset_index()
    bar_data['amount'] = bar_data['amount'].astype(float)

    return pie_data, bar_data

# Function to prompt GenAI for analysis
def promt_to_genai(df):
    pie_data, bar_data = prepare_data_for_chart(df)

    # Prepare data frame for GenAI prompt
    data_summary = f"""
    Pie Chart Data (Total Spending by Category):
    {pie_data.to_string(index=False)}

    Bar Chart Data (Quarterly Spending Trends):
    {bar_data.to_string(index=False)}

    Total Spending: ${df['amount'].sum():.2f}
    Date range: {df['transaction_date'].min()} to {df['transaction_date'].max()}
    """

    # Initialize GenAI client
    client = genai.Client(api_key=os.getenv("GEMINI_API_KEY"))

    # Generate content using GenAI
    response = client.models.generate_content(
        model="gemini-2.5-flash",
        contents=f"""Analyze this financial data and provide insights and sumarize your analysis to put into presentation slides.

                    {data_summary}

                    Please provide:
                    1. Key insights from the spending distribution (pie chart)
                    2. Notable trends across quarters (bar chart)""",
        # Đoạn promt này ông sửa như nào cho hợp là được
    )

    return response.text

if __name__ == "__main__":
    
    # df = get_data_from_supabase()
    df = get_data_from_local(os.getenv("LOCAL_FILE_PATH"))
    pie_data, bar_data = prepare_data_for_chart(df)

    analysis = promt_to_genai(df)
    print("Analysis from GenAI:")   
    print(analysis)

