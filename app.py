from flask import Flask, render_template, request
import matplotlib.pyplot as plt
import matplotlib.dates as mdates
from datetime import datetime, timedelta
import pandas as pd

app = Flask(__name__)

@app.route('/', methods=['GET', 'POST'])
def index():
    if request.method == 'POST':
        project_type = request.form['project_type']
        project_name = request.form['project_name']
        start_date_str = request.form['start_date']

        try:
            start_date = datetime.strptime(start_date_str, "%Y-%m-%d")
        except ValueError as e:
            return render_template('index.html', error=str(e))

        stages = get_stages(project_type)
        project_timeline, excel_filename = create_project_timeline(stages, project_name, start_date)

        return render_template('result.html', project_type=project_type, project_name=project_name,
                               start_date=start_date.strftime("%Y-%m-%d"), project_timeline=project_timeline,
                               excel_filename=excel_filename)

    return render_template('index.html')


def get_stages(project_type):
    minor_change_stages = [
        "Project kickoff",
        "Drawing and BOM",
        "QGC0",
        "Cable Sourcing",
        "Air gap analysis",
        "Customer C sample",
        "Customer approval for drawing",
        "D sample Production Release",
        "Cable PPAP",
        "QGC4",
        "SOP"
    ]
    adapt_project_stages = [
        "Adapt Project kickoff",
        "Design and mold flow analysis",
        "Drawing and BOM Approval",
        "Adapt QGC0",
        "A Sourcing",
        "Tooling kickoff",
        "C sample pre",
        "C sample DVT",
        "Tool transfer",
        "Tooling Trial",
        "MCA",
        "D sample and BOM",
        "D sample pre",
        "PVT",
        "SOP"
    ]
    return minor_change_stages if project_type == '1' else adapt_project_stages


def create_project_timeline(stages, project_name, start_date):
    project_timeline = []
    current_stage_start = start_date

    # Fixed durations for each stage in weeks
    fixed_durations = {
        "Project kickoff": 2,
        "Drawing and BOM": 4,
        "QGC0": 3,
        "Cable Sourcing": 5,
        "Air gap analysis": 2,
        "Customer C sample": 4,
        "Customer approval for drawing": 3,
        "D sample Production Release": 4,
        "Cable PPAP": 6,
        "QGC4": 2,
        "SOP": 3,
        "Adapt Project kickoff": 2,
        "Design and mold flow analysis": 3,
        "Drawing and BOM Approval": 4,
        "Adapt QGC0": 3,
        "A Sourcing": 5,
        "Tooling kickoff": 2,
        "C sample pre": 4,
        "C sample DVT": 3,
        "Tool transfer": 2,
        "Tooling Trial": 4,
        "MCA": 3,
        "D sample and BOM": 4,
        "D sample pre": 2,
        "PVT": 3
    }

    # Stage dependencies
    dependencies = {
        "Drawing and BOM": "Project kickoff",
        "QGC0": "Project kickoff",
        "Cable Sourcing": "Drawing and BOM",
        "Air gap analysis": "Drawing and BOM",
        "Customer C sample": "Cable Sourcing",
        "Customer approval for drawing": "Customer C sample",
        "D sample Production Release": "Customer approval for drawing",
        "Cable PPAP": "D sample Production Release",
        "QGC4": "Cable PPAP",
        "SOP": "QGC4",
        "Design and mold flow analysis": "Adapt Project kickoff",
        "Drawing and BOM Approval": "Design and mold flow analysis",
        "Adapt QGC0": "Adapt Project kickoff",
        "A Sourcing": "Drawing and BOM Approval",
        "Tooling kickoff": "Drawing and BOM Approval",
        "C sample pre": "Tooling kickoff",
        "C sample DVT": "C sample pre",
        "Tool transfer": "C sample pre",
        "Tooling Trial": "Tool transfer",
        "MCA": "Tooling Trial",
        "D sample and BOM": "MCA",
        "D sample pre": "D sample and BOM",
        "PVT": "D sample pre",
        "SOP": "PVT"
    }

    for stage_name in stages:
        if stage_name in dependencies:
            dependency_stage = dependencies[stage_name]
            dependency_stage_data = next((stage for stage in project_timeline if stage['Stage Name'] == dependency_stage), None)
            if dependency_stage_data:
                current_stage_start = dependency_stage_data['End Date']

        # Use fixed duration
        duration_weeks = fixed_durations.get(stage_name, 2)

        current_stage_end = current_stage_start + timedelta(weeks=duration_weeks)

        weeks_list = [current_stage_start + timedelta(weeks=i) for i in range(duration_weeks)]
        week_numbers = [week.isocalendar()[1] for week in weeks_list]

        project_timeline.append({
            'Stage Name': stage_name,
            'Start Date': current_stage_start,
            'End Date': current_stage_end,
            'Duration (weeks)': duration_weeks,
            'Weeks': week_numbers
        })

        current_stage_start = current_stage_end

    df = pd.DataFrame(project_timeline)

    # Generate the 'Weeks' column for all stages (list of week numbers)
    df['Weeks'] = df.apply(lambda row: f"Week {row['Start Date'].isocalendar()[1]} to Week {row['End Date'].isocalendar()[1]}", axis=1)

    excel_filename = "project_timeline.xlsx"
    df.to_excel(excel_filename, index=False)

    display_calendar(start_date, current_stage_end, project_timeline)

    return project_timeline, excel_filename


def display_calendar(start_date, end_date, project_timeline):
    fig, ax = plt.subplots(figsize=(12, 6))
    ax.xaxis.set_major_locator(mdates.MonthLocator())
    ax.xaxis.set_major_formatter(mdates.DateFormatter('%b %Y'))

    for stage in project_timeline:
        ax.plot([stage['Start Date'], stage['End Date']], [stage['Stage Name'], stage['Stage Name']], marker='o')

    ax.set_yticks(range(len(project_timeline)))
    ax.set_yticklabels([stage['Stage Name'] for stage in project_timeline])
    ax.set_title('Project Timeline')
    ax.set_xlabel('Timeline')
    ax.set_ylabel('Stages')
    plt.grid(True)

    plt.tight_layout()
    plt.savefig('static/calendar.png')


if __name__ == '__main__':
    app.run(debug=True)
