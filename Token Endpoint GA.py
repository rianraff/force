import requests
import json

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
    id_dict = {}
    if response.status_code == 200:
        json_response = json.loads(response.text)
        contents = json_response["data"]["content"]
        for content in contents:
            id_dict[content["description"]] = content["id"]
    else:
        print(f"Request failed with status code: {response.status_code}")
        print(f"Request failed with status code: {response.text}")

    return id_dict

def getDetailRequest(token, id_dict):
    document_dict = {}
    for cluster_id, id in id_dict.items():
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
            details_dict = {}
            for content in contents:
                details_dict[content["name"]] = content["minioId"]
            document_dict[cluster_id] = details_dict
        else:
            print(f"Request failed with status code: {response.status_code}")
            print(f"Request failed with status code: {response.text}")
    return document_dict

def main():
    token = getToken("amunawar@xl.co.id", "UntoeKF0rc3s#")
    id_dict = getRequestApproval(token, "14-03-2024", "14-03-2024")
    print(getDetailRequest(token, id_dict))


if __name__ == "__main__":
    main()