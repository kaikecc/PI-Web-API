# PI Web API Overview

This project demonstrates how to interact with the PI Web API to retrieve data and manage PI System resources. The following guide provides a comprehensive overview of the included code and its functionalities.

## Import Libraries

First, the necessary libraries are imported, including `requests` for making HTTP requests, `json` for handling JSON data, and `urllib.parse` for parsing URLs.

```python
import requests
from requests.auth import HTTPBasicAuth
import json
from urllib.parse import urlparse
```

## Base Code
A class PIWebAPI is defined to handle interactions with the PI Web API. It includes methods for initializing the connection and retrieving data.

### Class Initialization
The constructor method initializes the class with the endpoint URL, username, and password for authentication.

```python
def get_PiWebApi(self, custom_url=None):
    response_default = {"Links": {}, "Items": []}
    if not custom_url:
        print("No URL provided for the GET PI Web API data retrieval.")
        return None
    
    parsed_url = urlparse(f"{self.url_endpoint}/{custom_url}")
    if not all([parsed_url.scheme, parsed_url.netloc]):
        print(f"Invalid URL provided: {custom_url}")
        return None
    
    try:
        response = requests.get(f"{self.url_endpoint}{custom_url}", auth=HTTPBasicAuth(self.username, self.password), verify=False)
        if response.status_code == 200:
            return response.json()
        elif response.status_code == 204:
            print("204: Successful request but no content returned from the GET PI Web API.")
            return response_default
        else:
            print(f"Failed to retrieve data: {response.status_code} {response.reason}")
            return response_default
    except requests.exceptions.RequestException as e:
        print(f"Request failed: {e}")
        return response_default

```

## Usage Examples

### Retrieve Elements by Category Name

This example demonstrates how to filter elements by a specific category name using the get_PiWebApi method.

```python
category_name = "Equipment"
data = api.get_PiWebApi(f"elements/{webID_element}/elements?categoryName={category_name}")
print(data)
```

## Access Attributes in an Element

The following code shows how to access the attributes within a specified element.

```python
webID_attribute = "<F1DSUy1TdGF0ZXMgV2Vic2l0ZS9FbGVtZW50cw>"
data = api.get_PiWebApi(f"elements/{webID_attribute}/attributes")
print(data)
```

## Access Analyses in an Element

This example illustrates how to retrieve analyses associated with a specific element.

```python
webID_element = "<F1DSUy1TdGF0ZXMgV2Vic2l0ZS9FbGVtZW50cw>"
data = api.get_PiWebApi(f"elements/{webID_element}/analyses")
print(data)
```

## Conclusion

This project provides a basic framework for interacting with the PI Web API. It includes methods for initializing the API connection, retrieving elements by category, accessing attributes, and obtaining analyses. Modify and expand these examples to suit your specific requirements for PI System data management and retrieval.