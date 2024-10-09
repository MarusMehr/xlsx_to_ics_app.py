import pandas as pd
from ics import Calendar, Event
import streamlit as st
import datetime as dt
from io import BytesIO

def create_ics(df):
    cal = Calendar()
    for _, row in df.iterrows():
        event = Event()
        event.name = row['SUMMARY']
        event.begin = dt.datetime.strptime(str(row['DTSTART']), '%Y-%m-%d')
        event.end = dt.datetime.strptime(str(row['DTEND']), '%Y-%m-%d')
        event.location = row['LOCATION']
        event.description = f"Kontakt: {row['DESCRIPTION']}" if 'DESCRIPTION' in row else "Keine weiteren Informationen"
        cal.events.add(event)
    return cal

def generate_ics_file(cal):
    ics_file = BytesIO()
    ics_file.write(str(cal).encode('utf-8'))
    ics_file.seek(0)
    return ics_file

st.title("Excel to ICS Calendar Converter")

st.write("Upload an Excel file with the columns: `SUMMARY`, `DTSTART`, `DTEND`, `LOCATION`, `DESCRIPTION`")

uploaded_file = st.file_uploader("Choose an Excel file", type="xlsx")

if uploaded_file:
    df = pd.read_excel(uploaded_file)
    st.write("Preview of the data:")
    st.write(df.head())

    if all(col in df.columns for col in ["SUMMARY", "DTSTART", "DTEND", "LOCATION"]):
        cal = create_ics(df)
        ics_file = generate_ics_file(cal)
        
        st.download_button(
            label="Download ICS file",
            data=ics_file,
            file_name="calendar.ics",
            mime="text/calendar"
        )
    else:
        st.error("The Excel file must contain the columns: SUMMARY, DTSTART, DTEND, LOCATION, DESCRIPTION (optional)")
