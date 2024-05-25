# PIWebApi Client

This C# project provides a simple client for interacting with the PI Web API. It includes methods for retrieving and updating data, as well as testing the connection to the PI Web API server.

## Features

- **Get Data:** Fetch data from a specified PI Web API endpoint.


## Prerequisites

- .NET Core or .NET Framework
- NuGet Packages:
  - `Newtonsoft.Json`

## Environment Variables

Set the following environment variables for authentication:

- `PIWEBAPI_USERNAME`: Your PI Web API username.
- `PIWEBAPI_PASSWORD`: Your PI Web API password.

## Usage

### 1. Initialize the Client

```csharp
string urlEndpoint = "https://servername/piwebapi/";
PIWebApi piWebApi = new PIWebApi(urlEndpoint);
```

### 2. Get Data from PI Web API

```csharp
string webId_element = "E0gASZy4oKQ9kiBiZJTW7eugwQVgAAAABQVJNSUQxMDAwMA";
string customUrl = $"elements/{webId_element})";
dynamic result = await piWebApi.GetPiWebApiAsync(customUrl);
Console.WriteLine(result);
```


### Code Overview
PIWebApi Class
This class contains methods for interacting with the PI Web API.

* Constructor:

```csharp
public PIWebApi(string urlEndpoint)
```

* Get Data:
  
```csharp
public async Task<dynamic> GetPiWebApiAsync(string customUrl)
```



### License
This project is licensed under the MIT License. See the LICENSE file for more details.

### Acknowledgments
Newtonsoft.Json for JSON serialization and deserialization.
PI Web API documentation for providing API endpoints and usage guidelines.

### Contributing

Contributions are welcome! Please open an issue or submit a pull request for any improvements or bug fixes.