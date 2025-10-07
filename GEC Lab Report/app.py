from flask import Flask, render_template, request, send_file
import pandas as pd
import pyodbc
from io import BytesIO
from datetime import datetime

app = Flask(__name__)

@app.route('/', methods=['GET', 'POST'])
def index():
    if request.method == 'POST':
        start = request.form['start']
        end = request.form['end']

        try:
            # Convert to datetime format
            start_dt = datetime.strptime(start, "%Y-%m-%dT%H:%M")
            end_dt = datetime.strptime(end, "%Y-%m-%dT%H:%M")

            # Connect to SQL Server
            conn = pyodbc.connect(
                "DRIVER={SQL Server};"
                "SERVER=10.107.34.16;"
                "DATABASE=ConnectedFactory;"
                "UID=sa;"
                "PWD=Noida@#geclab12525"
            )

            query = f"""
                    SELECT
                        A.DateTime,
                        A.Pow_Ana_SinglePhase1_Voltage AS Voltage,
                        A.Pow_Ana_SinglePhase1_Current AS Curren,
                        A.Pow_Ana_SinglePhase1_KVA AS KVA,
                        A.Pow_Ana_SinglePhase1_KW AS KW,
                        A.Pow_Ana_SinglePhase1_Power_Factor AS PF,
                        A.Pow_Ana_SinglePhase1_Total_KW AS Total,
                        A.Pow_Ana_SinglePhase1_Frequency AS Fr,
                        B.Sensor1,
                        B.Sensor2,
                        B.Sensor3,
                        B.Sensor4,
                        B.Sensor5,
                        B.Sensor6,
                        B.Sensor7,
                        B.Sensor8,
                        C."Pow_Ana_ThreePhase1_Urms&#00931;" AS Urms,
                        C. "Pow_Ana_ThreePhase1_Irms&#00931;" AS Irms,
                        C."Pow_Ana_ThreePhase1_Lambda&#00931;" AS Lambda,
                        C."Pow_Ana_ThreePhase1_P&#00931;" AS P,
                        C."Pow_Ana_ThreePhase1_S&#00931;" AS S,
                        C."Pow_Ana_ThreePhase1_fU1" AS fU1,
                        C."Pow_Ana_ThreePhase1_fI1" AS fI1
                        FROM 
                        SINGLEPHASEANALYSER A
                        JOIN 
                        MEGHDOOT B 
                            ON A.DateTime = B.DateTime
                        JOIN 
                        THREEPHASEANALYSER C 
                            ON A.DateTime = C.DateTime
                        WHERE 
                        A.DateTime BETWEEN '{start_dt}' AND '{end_dt}'
                        ORDER BY 
                        A.DateTime DESC;
                            """

            df = pd.read_sql(query, conn)
            conn.close()

            # Generate Excel
            output = BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                df.to_excel(writer, index=False, sheet_name='Report')
            output.seek(0)

            return send_file(
                output,
                download_name="GEC_Lab_Phase_Analyser_and_Meghdoot_Report.xlsx",
                as_attachment=True,
                mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

        except Exception as e:
            return f"Error: {e}"

    return render_template('index.html')

# if __name__ == '__main__':
#     print("Flask app is running. Open http://127.0.0.1:5000/ in your browser.")
#     app.run(debug=True)
if __name__ == '__main__':
    try:
        print("Starting Flask app...")
        app.run(debug=True)
    except Exception as e:
        print(f"Error starting app: {e}")
