import requests
import json
import os
import psycopg2
from dotenv import load_dotenv
from flask import Flask, request, jsonify


#----------------------
# Get Approve API
#----------------------

def getToken(username, password):
    url = 'https://www.getapprove.xl.co.id/api/v2/oauth/token'
    headers = {
        'Authorization': 'Basic YW11bmF3YXJAeGwuY28uaWQ6VW50b2VLRjByYzNzIw==',
        'Content-Type': 'application/x-www-form-urlencoded',
        'Cookie': '__cf_bm=wwUTWPkA3DzXroYd_yVvI66ZZqBFuLIrKNtaYTH4DNw-1710741277-1.0.1.1-wXWb3st3LLiJ08z6.F7.XW4ysQeKf4bDpuf9blvtybB3Kwnn5zCHxyTeIjNVtkAZnr.RHNvbj6HMaQuo6j_voQ'
    }

    data = {
        'grant_type': 'password',
        'username': username,
        'password': password
    }

    # Make a POST request to the API endpoint with data and headers
    response = requests.post(url, data=data, headers=headers)

    # Check if the request was successful (status code 200)
    if response.status_code == 200:
        # Print the response content (JSON, XML, or text)
        # print(response.json())  # If the response is JSON
        json_response = json.loads(response.text)  # If the response is plain text
        token = json_response["access_token"]
        # print(response.content)  # If you want the raw content
    else:
        # If the request was unsuccessful, print the status code
        print(f"Request failed with status code: {response.status_code}")

    return token

def getRequestApproval(token, startCreatedDate, endCreatedDate):
    url = 'https://www.getapprove.xl.co.id/api/v2/bpmn/requestDocument/dashboard-admin'
    params = {
        'name': 'RFS Homepass',
        'startCreatedDate': startCreatedDate,
        'endCreatedDate': endCreatedDate,
        'sort': 'createdDate',
        'asc': '0'
    }

    headers = {
        'Content-Type': 'application/json',
        'Authorization': token,
        'Cookie': '__cf_bm=cXzQEE9tr05sjFIWOFtoB033onNNsMjiuuOfs92P2Ew-1710736596-1.0.1.1-e0AI1WbIvwm7.V5KjVXQjilLFh_XX0oxEexOmx8WoYOvgI0RGcsP_bSTwNCTSdRIloXrSDhbZSuiku092ohiSw'
    }

    base_dir = os.path.join(os.getcwd(), "Input")

    response = requests.get(url, params=params, headers=headers)
    if response.status_code == 200:
        json_response = json.loads(response.text)
        contents = json_response["data"]["content"]
        net_contents = []
        for content in contents:
            net_contents.append(content["description"])
        net_contents = list(set(net_contents))
        for content in net_contents:
            insert_data_to_db(conn, content, 'FALSE', '', 'FALSE')
        for content in contents:
            if os.path.exists(os.path.join(base_dir, content["description"])):
                os.replace(os.path.join(base_dir, content["description"]), os.path.join(base_dir, content["description"]))
            else:
                os.makedirs(os.path.join(base_dir, content["description"]))
            getDetailRequest(token, content["id"], content["description"])
            update_path_in_db(conn, content["description"], os.path.join(base_dir, content["description"]))
    else:
        print(f"Request failed with status code: {response.status_code}")
        print(f"Request failed with status code: {response.text}")

    

def getDetailRequest(token, id, cluster_id):
    url = f'https://www.getapprove.xl.co.id/api/v2/bpmn/requestDocument/{id}'

    headers = {
        'User-ID': 'idsta.harynr@xl.co.id',
        'Authorization': token,
        'Cookie': '__cf_bm=hPi1UnexJtcsbaxvM8plQ6MAa0Hy99GzwXYCtHbrRTQ-1710737810-1.0.1.1-RFUOs6hQmeTyhbffEKf3IlWqBCabsL1I6jrqIrIXRyotr4Z17FvkaPUblHmwYf4HxOpjnIbn0zfmF3RhVrf.dA'
    }
    response = requests.get(url, headers=headers)

    if response.status_code == 200:
        json_response = json.loads(response.text)
        contents = json_response["data"]["object"]["fileDocuments"]
        for content in contents:
            downloadReqApproval(token, cluster_id, content["name"], content["minioId"])
    else:
        print(f"Request failed with status code: {response.status_code}")
        print(f"Request failed with status code: {response.text}")

def downloadReqApproval(token, cluster_id, file_name, minioId):
    url = f'https://www.getapprove.xl.co.id/api/v2/bpmn/requestDocument/download-file?minioId={minioId}'
    # Replace with the URL of the file you want to download
    filename = os.path.join(cluster_id, file_name)  # Specify the filename for the downloaded file
    headers = {
        'User-ID': 'idsta.harynr@xl.co.id',
        'Authorization': token,
    }
    base_dir = os.path.join(os.getcwd(), "Input")
    # Send a GET request to the URL
    response = requests.get(url, headers=headers)

    # Check if the request was successful (status code 200)
    if response.status_code == 200:
        # Open the file in binary write mode
        with open(os.path.join(base_dir,filename), 'wb') as f:
            # Write the content of the response to the file
            f.write(response.content)
        print(f"File '{filename}' downloaded successfully.")
    else:
        print(f"Failed to download file. Status code: {response.status_code}")

#----------------------
# Force Database
#----------------------
    
# Load environment variables from .env file
load_dotenv()

# Create a PostgreSQL connection
conn = psycopg2.connect(
    dbname=os.getenv("POSTGRES_DB"),
    user=os.getenv("POSTGRES_USER"),
    password=os.getenv("POSTGRES_PASSWORD"),
    host=os.getenv("POSTGRES_HOST"),
    port=os.getenv("POSTGRES_PORT")
)

# Initialize Flask app
app = Flask(__name__)

# Route to insert data into the database
@app.route('/insert_data', methods=['POST'])
def insert_data():
    data = request.get_json()
    cluster_id = data.get('cluster_id')
    processed = data.get('processed')
    file_path = data.get('file_path')
    revise = data.get('revise')

    if not cluster_id or not processed or not revise:
        return jsonify({'message': 'Missing required parameters'}), 400

    insert_data_to_db(conn, cluster_id, processed, file_path, revise)
    return jsonify({'message': 'Data inserted successfully'}), 200

def insert_data_to_db(conn, cluster_id, processed, file_path, revise):
    # Create a new cursor
    cur = conn.cursor()

    # Execute the INSERT statement
    cur.execute(
        "INSERT INTO logging (cluster_id, processed, file_path, revise) VALUES (%s, %s, %s, %s)",
        (cluster_id, processed, file_path, revise)
    )

    # Commit the changes to the database
    conn.commit()

    # Close the cursor
    cur.close()

# Route to get list of clusters
@app.route('/clusters', methods=['GET'])
def clusters():
    return getClusterList()

def getClusterList():
    try:
        # Create a new cursor
        cur = conn.cursor()

        # Execute a query to select all clusters from the logging table
        cur.execute("SELECT DISTINCT cluster_id FROM logging")

        # Fetch all rows from the result set
        rows = cur.fetchall()

        # Close the cursor
        cur.close()

        # Return the list of clusters as a JSON response
        return jsonify({"clusters": [row[0] for row in rows]})

    except psycopg2.Error as e:
        # Handle any errors that occur during database operation
        print(f"Error retrieving cluster list: {e}")
        return jsonify({"error": "Failed to retrieve cluster list"}), 500
    
@app.route('/update_path', methods=['POST'])
def update_path():
    data = request.json
    cluster_id = data.get('cluster_id')
    file_path = data.get('new_path')

    if not cluster_id or not file_path:
        return jsonify({'error': 'Missing cluster_id or new_path parameter'}), 400

    update_path_in_db(conn, cluster_id, file_path)

def update_path_in_db(conn, cluster_id, file_path):
    # Create a new cursor
    cur = conn.cursor()

    # Execute the UPDATE statement
    cur.execute(
        "UPDATE logging SET file_path = %s WHERE cluster_id = %s",
        (file_path, cluster_id)
    )

    # Commit the changes to the database
    conn.commit()

    # Close the cursor
    cur.close()
    
def main():
    token = getToken("amunawar@xl.co.id", "UntoeKF0rc3s#")
    getRequestApproval(token, "14-03-2024", "14-03-2024")


if __name__ == '__main__':
    # app.run(debug=True)
    main()