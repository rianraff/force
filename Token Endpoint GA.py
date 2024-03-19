import requests
import json

def getToken():
    url = 'https://www.getapprove.xl.co.id/api/v2/oauth/token'
    headers = {
        'Authorization': 'Basic YW11bmF3YXJAeGwuY28uaWQ6VW50b2VLRjByYzNzIw==',
        'Content-Type': 'application/x-www-form-urlencoded',
        'Cookie': '__cf_bm=wwUTWPkA3DzXroYd_yVvI66ZZqBFuLIrKNtaYTH4DNw-1710741277-1.0.1.1-wXWb3st3LLiJ08z6.F7.XW4ysQeKf4bDpuf9blvtybB3Kwnn5zCHxyTeIjNVtkAZnr.RHNvbj6HMaQuo6j_voQ'
    }

    data = {
        'grant_type': 'password',
        'username': 'amunawar@xl.co.id',
        'password': 'UntoeKF0rc3s#'
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

def getRequestApproval(token):
    url = 'https://www.getapprove.xl.co.id/api/v2/bpmn/requestDocument/dashboard-admin'
    params = {
        'name': 'RFS Homepass',
        'startCreatedDate': '14-03-2024',
        'endCreatedDate': '14-03-2024',
        'sort': 'createdDate',
        'asc': '0'
    }

    headers = {
        'Content-Type': 'application/json',
        'Authorization': token,
        'Cookie': '__cf_bm=wwUTWPkA3DzXroYd_yVvI66ZZqBFuLIrKNtaYTH4DNw-1710741277-1.0.1.1-wXWb3st3LLiJ08z6.F7.XW4ysQeKf4bDpuf9blvtybB3Kwnn5zCHxyTeIjNVtkAZnr.RHNvbj6HMaQuo6j_voQ'
    }

    response = requests.get(url, params=params, headers=headers)

    if response.status_code == 200:
        print(response.json())
    else:
        print(f"Request failed with status code: {response.status_code}")
        print(f"Request failed with status code: {response.text}")


def main():
    token = getToken()
    getRequestApproval(token)

if __name__ == "__main__":
    main()