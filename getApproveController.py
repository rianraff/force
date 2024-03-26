import datetime
import requests
import json
import os
import psycopg2
from dotenv import load_dotenv
from flask import Flask, request, jsonify


#----------------------
# Get Approve API
#----------------------
force_base_url = 'http://localhost:5000'

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
        response = requests.get(url)

        # force get all cluster ids
        url = f'{force_base_url}/get_all_cluster_ids'  
        response = requests.get(url)
        if response.status_code == 200:
            cluster_ids = response.json()
        else:
            print('Failed to get cluster IDs. Status code:', response.status_code)
            
        for content in net_contents:
            if content not in cluster_ids:
                # force insert api
                url = f"{force_base_url}/insert_data"
                data = {
                    'cluster_id': content,
                    'processed': 'FALSE',
                    'file_path': '',
                    'revise': 'FALSE'
                }
                headers = {'Content-Type': 'application/json'}
                response = requests.post(url, headers=headers, data=json.dumps(data))

            if response.status_code == 200:
                print('Data inserted successfully.')
            else:
                print('Failed to insert data. Status code:', response.status_code)

        for content in contents:
            if os.path.exists(os.path.join(base_dir, content["description"])):
                os.replace(os.path.join(base_dir, content["description"]), os.path.join(base_dir, content["description"]))
            else:
                os.makedirs(os.path.join(base_dir, content["description"]))
            getDetailRequest(token, content["id"], content["description"])
            
            # force update api
            url = f"{force_base_url}/update_path"
            data = {
                'cluster_id': content["description"],
                'file_path': os.path.join(base_dir, content["description"])
            }
            headers = {'Content-Type': 'application/json'}

            response = requests.post(url, headers=headers, data=json.dumps(data))

            if response.status_code == 200:
                print('Path updated successfully.')
            else:
                print('Failed to update path. Status code:', response.status_code)
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

def main():
    token = getToken("amunawar@xl.co.id", "UntoeKF0rc3s#")
    today_time = datetime.date.today().strftime('%d-%m-%Y')
    getRequestApproval(token, today_time, today_time)

if __name__ == '__main__':
    # app.run(debug=True)
    main()