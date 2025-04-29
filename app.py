import pandas as pd
import streamlit as st
import time
from datetime import datetime, timedelta

# Set up the title and description of the app
st.title("Student Appointment Organizer")
st.write("Upload your appointment schedule in Excel format to organize them into In-Person and Zoom meetings.")

# File uploader for user to upload the Excel file
uploaded_file = st.file_uploader("Upload your Excel file", type=["xlsx"])

if uploaded_file is not None:
    # Load the Excel file into a pandas DataFrame
    df = pd.read_excel(uploaded_file)

    # Convert 'Scheduled Start Time' to datetime if it's not already in datetime format
    df['Scheduled Start Time'] = pd.to_datetime(df['Scheduled Start Time'], errors='coerce')

    # Get the current date
    current_date = datetime.now()

    # Filter the data for the past month, half year, and full year
    one_month_ago = current_date - timedelta(days=30)
    six_months_ago = current_date - timedelta(days=180)
    one_year_ago = current_date - timedelta(days=365)

    df_month = df[df['Scheduled Start Time'] >= one_month_ago]
    df_half_year = df[df['Scheduled Start Time'] >= six_months_ago]
    df_year = df[df['Scheduled Start Time'] >= one_year_ago]

    # Function to get the host with the most meetings
    def most_meetings_by_host(df):
        host_counts = df['Host'].value_counts()
        return host_counts.idxmax(), host_counts.max()

    # Function to get the count of virtual and in-person meetings
    def virtual_vs_in_person(df):
        virtual_count = df[df['Location'] == 'Zoom'].shape[0]
        in_person_count = df[df['Location'] == 'College of Engineering'].shape[0]
        return virtual_count, in_person_count

    # Function to get the student with the most meetings
    def most_meetings_by_student(df):
        student_counts = df.groupby(['First Name', 'Last Name']).size()
        most_active_student = student_counts.idxmax()
        return most_active_student, student_counts.max()

    # Generate the insights for the past month
    host_month, host_count_month = most_meetings_by_host(df_month)
    virtual_month, in_person_month = virtual_vs_in_person(df_month)
    student_month, student_count_month = most_meetings_by_student(df_month)

    # Insights for the past half year
    host_half_year, host_count_half_year = most_meetings_by_host(df_half_year)
    virtual_half_year, in_person_half_year = virtual_vs_in_person(df_half_year)
    student_half_year, student_count_half_year = most_meetings_by_student(df_half_year)

    # Insights for the past year
    host_year, host_count_year = most_meetings_by_host(df_year)
    virtual_year, in_person_year = virtual_vs_in_person(df_year)
    student_year, student_count_year = most_meetings_by_student(df_year)

    # Display insights
    st.subheader("Insights Over the Past Month:")
    st.write(f"Host with the most meetings: {host_month} ({host_count_month} meetings)")
    st.write(f"Virtual meetings: {virtual_month}, In-Person meetings: {in_person_month}")
    st.write(f"Most active student: {student_month} ({student_count_month} meetings)")

    st.subheader("Insights Over the Past Half Year:")
    st.write(f"Host with the most meetings: {host_half_year} ({host_count_half_year} meetings)")
    st.write(f"Virtual meetings: {virtual_half_year}, In-Person meetings: {in_person_half_year}")
    st.write(f"Most active student: {student_half_year} ({student_count_half_year} meetings)")

    st.subheader("Insights Over the Past Year:")
    st.write(f"Host with the most meetings: {host_year} ({host_count_year} meetings)")
    st.write(f"Virtual meetings: {virtual_year}, In-Person meetings: {in_person_year}")
    st.write(f"Most active student: {student_year} ({student_count_year} meetings)")

    # Ask if the user wants to download a summary sheet with all these insights
    download_combined = st.radio(
        "Would you like to download an Excel sheet with all data insights (separated into sheets)?",
        ("Yes", "No")
    )

    if download_combined == "Yes":
        # Create a combined Excel file with separate sheets for each time period
        with pd.ExcelWriter("combined_insights_appointments.xlsx") as writer:
            # Create a DataFrame to store the summary for each period
            summary_data = {
                "Time Period": ["Past Month", "Past Half Year", "Past Year"],
                "Most Active Host": [host_month, host_half_year, host_year],
                "Host Meeting Count": [host_count_month, host_count_half_year, host_count_year],
                "Virtual Meetings": [virtual_month, virtual_half_year, virtual_year],
                "In-Person Meetings": [in_person_month, in_person_half_year, in_person_year],
                "Most Active Student": [student_month, student_half_year, student_year],
                "Student Meeting Count": [student_count_month, student_count_half_year, student_count_year]
            }
            summary_df = pd.DataFrame(summary_data)

            # Write summary to 'Summary' sheet
            summary_df.to_excel(writer, sheet_name="Summary Insights", index=False)

            # Write the detailed data for each time period
            df_month.to_excel(writer, sheet_name="Month Data", index=False)
            df_half_year.to_excel(writer, sheet_name="Half Year Data", index=False)
            df_year.to_excel(writer, sheet_name="Year Data", index=False)

        # Provide the file to be downloaded
        with open("combined_insights_appointments.xlsx", "rb") as f:
            st.download_button(
                label="Download Combined Insights Sheet",
                data=f,
                file_name="combined_insights_appointments.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

    elif download_combined == "No":
        st.write("No problem, I'm here when you need!")
