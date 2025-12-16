from supabase import create_client
import google.generativeai as genai
from dotenv import load_dotenv
from pptx import Presentation
from pptx.chart.data import CategoryChartData
from tenacity import retry, stop_after_attempt, wait_exponential, retry_if_exception_type
from google.api_core import exceptions as google_exceptions
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
@retry(
    retry=retry_if_exception_type(google_exceptions.ServerError),
    wait=wait_exponential(multiplier=1, min=4, max=10),
    stop=stop_after_attempt(3)
)
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
    genai.configure(api_key=os.getenv("GEMINI_API_KEY"))

    # Set up the model
    model = genai.GenerativeModel(model_name="gemini-2.5-flash")

    # Generate content using GenAI
    prompt = f"""Analyze this financial data and provide insights and sumarize your analysis to put into presentation slides.

                    {data_summary}

                    Please provide:
                    1. Key insights from the spending distribution (pie chart)
                    2. Notable trends across quarters (bar chart)"""
    response = model.generate_content(prompt)

    return response.text

# Function to update charts in a PowerPoint presentation
def update_presentation(pie_data, bar_data, template_path, output_path):
    """
    Updates charts in a PowerPoint presentation from a template.
    - Slide 1 is expected to have a Pie Chart.
    - Slide 2 is expected to have a Clustered Column Chart.
    """
    prs = Presentation(template_path)

    # --- Update Pie Chart on Slide 1 ---
    pie_slide = prs.slides[0]
    # Assume the chart is the first shape of type 'chart' on the slide
    pie_chart_shape = [s for s in pie_slide.shapes if s.has_chart][0]
    pie_chart = pie_chart_shape.chart

    chart_data = CategoryChartData()
    chart_data.categories = pie_data['category']
    chart_data.add_series('Amount', pie_data['amount'])

    pie_chart.replace_data(chart_data)

    # Set chart title for Pie Chart
    if pie_chart.has_title:
        pie_chart.chart_title.text_frame.text = "Pie Chart"


    # --- Update Bar Chart on Slide 2 ---
    bar_slide = prs.slides[1]
    # Assume the chart is the first shape of type 'chart' on the slide
    bar_chart_shape = [s for s in bar_slide.shapes if s.has_chart][0]
    bar_chart = bar_chart_shape.chart

    # Pivot data to have quarters as categories and spending categories as series
    pivoted_bar_data = bar_data.pivot(index='quarter', columns='category', values='amount').fillna(0)

    chart_data = CategoryChartData()
    chart_data.categories = pivoted_bar_data.index.astype(str) # Quarters

    for col in pivoted_bar_data.columns:
        chart_data.add_series(col, pivoted_bar_data[col]) # Add each spending category as a series

    bar_chart.replace_data(chart_data)
    
    # Set chart title for Bar Chart
    if bar_chart.has_title:
        bar_chart.chart_title.text_frame.text = "Bar Chart"

    prs.save(output_path)
    print(f"Presentation saved to {output_path}")

if __name__ == "__main__":
    
    # df = get_data_from_supabase()
    # Construct the path to the data file relative to the script's location
    script_dir = os.path.dirname(__file__)
    file_path = os.path.join(script_dir, '..', 'data', 'xlsx', 'financial_data.xlsx')
    df = get_data_from_local(file_path)
    pie_data, bar_data = prepare_data_for_chart(df)

    # Define presentation paths
    template_presentation_path = os.path.join(script_dir, '..', 'data', 'power-point', 'template.pptx')
    output_presentation_path = os.path.join(script_dir, '..', 'data', 'power-point', 'financial_report.pptx')

    # Update the presentation with new data
    update_presentation(pie_data, bar_data, template_presentation_path, output_presentation_path)

    analysis = promt_to_genai(df)
    print("Analysis from GenAI:")   
    print(analysis)