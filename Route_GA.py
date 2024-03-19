import requests
import json
import os

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

    response = requests.get(url, params=params, headers=headers)
    if response.status_code == 200:
        json_response = json.loads(response.text)
        contents = json_response["data"]["content"]
        for content in contents:
            os.makedirs(os.path.join(os.getcwd(), content["description"]))
            getDetailRequest(token, content["id"], content["description"])
            
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
    file_path = os.getcwd()
    # Send a GET request to the URL
    response = requests.get(url, headers=headers)

    # Check if the request was successful (status code 200)
    if response.status_code == 200:
        # Open the file in binary write mode
        with open(os.path.join(file_path,filename), 'wb') as f:
            # Write the content of the response to the file
            f.write(response.content)
        print(f"File '{filename}' downloaded successfully.")
    else:
        print(f"Failed to download file. Status code: {response.status_code}")

def main():
    token = getToken("amunawar@xl.co.id", "UntoeKF0rc3s#")
    getRequestApproval(token, "14-03-2024", "14-03-2024")
    

if __name__ == "__main__":
    main()