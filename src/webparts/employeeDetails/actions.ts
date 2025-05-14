import { WebPartContext } from "@microsoft/sp-webpart-base";
import { SPHttpClient, SPHttpClientResponse } from "@microsoft/sp-http";
import { IEmployee } from "./components/IEmployee";

const siteUrl = "https://multithreadgroup.sharepoint.com/sites/Playground";

//fetchEmployees function
// This function fetches employee data from the SharePoint list named 'EmployeeDetails'.
export async function fetchEmployees(
  context: WebPartContext
): Promise<IEmployee[]> {
  const url = `${siteUrl}/_api/web/lists/GetByTitle('EmployeeDetails')/items`;

  try {
    const response: SPHttpClientResponse = await context.spHttpClient.get(
      url,
      SPHttpClient.configurations.v1
    );

    if (!response.ok) {
      const text = await response.text();
      throw new Error(`HTTP ${response.status}: ${text}`);
    }

    const data = await response.json();
    return data.value as IEmployee[];
  } catch (error) {
    console.error("fetchEmployees error:", error);
    throw error;
  }
}

//addEmployee function
// This function adds a new employee to the SharePoint list named 'EmployeeDetails'.
export async function addEmployee(
  context: WebPartContext,
  employee: Partial<IEmployee>
): Promise<void> {
  const url = `${siteUrl}/_api/web/lists/GetByTitle('EmployeeDetails')/items`;

  const body: string = JSON.stringify({
    Title: employee.Name,
    Name: employee.Name,
    HireDate: employee.HireDate,
    JobDescription: employee.JobDescription,
  });

  console.log("addEmployee body:", body);

  try {
    const response: SPHttpClientResponse = await context.spHttpClient.post(
      url,
      SPHttpClient.configurations.v1,
      {
        headers: {
          Accept: "application/json; odata.metadata=minimal",
          "Content-Type": "application/json; charset=utf-8",
        },
        body,
      }
    );

    if (!response.ok) {
      const text = await response.text();
      throw new Error(`HTTP ${response.status}: ${text}`);
    }

    console.log("Employee added successfully.");
    alert("Employee added successfully");
  } catch (error) {
    console.error("addEmployee error:", error);
    throw error;
  }
}

//updateEmployee function
// This function updates an existing employee in the SharePoint list named 'EmployeeDetails'.
export async function updateEmployee(
  context: WebPartContext,
  id: number,
  employee: Partial<IEmployee>
): Promise<void> {
  const url = `${siteUrl}/_api/web/lists/GetByTitle('EmployeeDetails')/items(${id})`;
  const body = JSON.stringify({
    Title: employee.Name,
    Name: employee.Name,
    HireDate: employee.HireDate,
    JobDescription: employee.JobDescription,
  });

  try {
    const response: SPHttpClientResponse = await context.spHttpClient.post(
      url,
      SPHttpClient.configurations.v1,
      {
        headers: {
          Accept: "application/json; odata.metadata=minimal",
          "Content-Type": "application/json; charset=utf-8",
          "IF-MATCH": "*", // overwrite regardless of current ETag
          "X-HTTP-Method": "MERGE", // instruct SharePoint to merge/update
        },
        body,
      }
    );

    if (!response.ok) {
      const text = await response.text();
      throw new Error(`Update failed (HTTP ${response.status}): ${text}`);
    }

    console.log(`Employee ${id} updated successfully.`);
  } catch (err) {
    console.error("updateEmployee error:", err);
    throw err;
  }
}

//deleteEmployee function
// This function deletes an employee from the SharePoint list named 'EmployeeDetails'.
export async function deleteEmployee(
  context: WebPartContext,
  id: number
): Promise<void> {
  const url = `${siteUrl}/_api/web/lists/GetByTitle('EmployeeDetails')/items(${id})`;

  try {
    const response: SPHttpClientResponse = await context.spHttpClient.post(
      url,
      SPHttpClient.configurations.v1,
      {
        headers: {
          "IF-MATCH": "*",
          "X-HTTP-Method": "DELETE",
        },
      }
    );

    if (!response.ok) {
      const text = await response.text();
      throw new Error(`Delete failed (HTTP ${response.status}): ${text}`);
    }

    console.log(`Employee ${id} deleted successfully.`);
  } catch (err) {
    console.error("deleteEmployee error:", err);
    throw err;
  }
}
