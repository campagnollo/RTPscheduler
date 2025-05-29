Automated Team Scheduling Tool with Excel Output

RTPscheduler is a Python-based tool designed to automate the creation of team schedules and export them into
 Microsoft Excel format. This utility streamlines the scheduling process, making it efficient and error-free.

📋 Features
Automated Scheduling: Generates team schedules based on predefined parameters.

Excel Export: Outputs the schedule in .xlsx format for easy sharing and editing.

Customizable Inputs: Accepts input data from XML files, allowing for flexible scheduling configurations.

🚀 Getting Started
Prerequisites
Python 3.x installed on your system.

Required Python packages listed in requirements.txt.

Installation
Clone the Repository

bash
Copy
Edit
git clone https://github.com/campagnollo/RTPscheduler.git
cd RTPscheduler
Install Dependencies

bash
Copy
Edit
pip install -r requirements.txt

🖥️ Usage
Prepare Input Data

Ensure you have the necessary input files, such as MESS_list.xml, configured with your team's scheduling information.

Run the Script

Execute the main Python script to generate the schedule:

bash
Copy
Edit
python main.py
Access the Output

Upon successful execution, the script will produce an Excel file (e.g., 2022-10-10.xlsx) containing the generated schedule.

📄 Example
Given an input XML file with team member details and availability, running the script will output an Excel file
 structured as follows:

Name	Date	Shift Time
John Doe	2022-10-10	09:00-17:00
Jane Smith	2022-10-10	10:00-18:00

This format facilitates easy distribution and further editing if necessary.

🧑‍💻 Author
Eric L. Moore
DevOps & Cloud Infrastructure Engineer
GitHub Profile

📜 License
This project is licensed under the MIT License.