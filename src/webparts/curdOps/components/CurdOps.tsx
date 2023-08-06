import * as React from "react";
import { ISPHttpClientOptions, SPHttpClientResponse } from "@microsoft/sp-http";
import ContextService from "../loc/ContextService";
import { SPHttpClient } from "@microsoft/sp-http";
import "./CrudOOps.scss";

import CheckTheme from "./CheckTheme";
const CurdOps = () => {
  const [data, setData] = React.useState({
    username: "",
    pwd: "",
  });
  const [listData, setListData] = React.useState<any>([]);
  // const fieldRef = React.useRef(null);
  const handelChange = (event: any) => {
    const { name, value }: any = event.target;

    setData((prevProps) => ({
      ...prevProps,
      [name]: value,
    }));
  };
  React.useEffect(() => {
    getData();
  }, [listData]);
  const handleClick = async (event: any) => {
    event.preventDefault();

    data.username.trim() != "" && data.pwd.trim() != ""
      ? await addData(data.username, data.pwd)
      : null;
    console.log("Submited -:- successfully");
    setData({
      username: "",
      pwd: "",
    });
    console.log("data fetched...");
  };

  const addData = (username: any, password: any) => {
    // Define the URL of the SharePoint site and the list name
    const webUrl = ContextService.GetUrl();
    const listName = "RohitFirstListTesting";

    // Define the new item data as an object
    const newItemData = {
      UserName: `${username}`,
      Password: `${password}`,
    };

    // Prepare the request options
    const requestHeaders: Headers = new Headers();
    requestHeaders.append("Content-type", "application/json");
    const requestOptions: ISPHttpClientOptions = {
      headers: requestHeaders,
      body: JSON.stringify(newItemData),
    };
    // Define the endpoint URL for adding a new item to the list
    const endpointUrl = `${webUrl}/_api/web/lists/getbytitle('${listName}')/items`;

    // Send the POST request to create the new item
    ContextService.GetSPContext()
      .post(endpointUrl, SPHttpClient.configurations.v1, requestOptions)
      .then((response: SPHttpClientResponse) => {
        if (response.ok) {
          console.log("New item created successfully");
        } else {
          console.log("Error creating new item: " + response.statusText);
        }
      })
      .catch((error: any) => {
        console.log("Error creating new item: " + error);
      });
  };
  const getData = () => {
    const listname = "RohitFirstListTesting";
    const url = `${ContextService.GetUrl()}/_api/web/lists/getbytitle('${listname}')/items`;
    ContextService.GetSPContext()
      .get(url, SPHttpClient.configurations.v1)
      .then((response) => {
        if (response.ok) {
          response.json().then((data: any) => {
            // listData.push(data.value);
            setListData(data.value);
            // setListData((prev: any) => {
            //   prev;
            // });
          });
        } else {
          console.error("an error occured while fetching data");
        }
      });
  };
  const UpdateChange = (Id: any, Uname: any, Pass: any) => {
    // Define the URL of the SharePoint site and the list name
    const webUrl = ContextService.GetUrl();
    const listName = "RohitFirstListTesting";

    // Define the new item data as an object
    const newItemData = {
      UserName: Uname,
      Password: Pass,
    };

    const requestOptions: ISPHttpClientOptions = {
      headers: {
        Accept: "application/json;odata=nometadata",
        "Content-type": "application/json;odata=nometadata",
        "odata-version": "",
        "IF-MATCH": "*",
        "X-HTTP-Method": "MERGE",
      },
      body: JSON.stringify(newItemData),
    };
    // Define the endpoint URL for adding a new item to the list
    const endpointUrl = `${webUrl}/_api/web/lists/getbytitle('${listName}')/items(${Id})`;

    // Send the POST request to create the new item
    ContextService.GetSPContext()
      .post(endpointUrl, SPHttpClient.configurations.v1, requestOptions)
      .then((response: SPHttpClientResponse) => {
        if (response.ok) {
          console.log("item Updated successfully");
        } else {
          console.log("Error Updating new item: " + response.statusText);
        }
      })
      .catch((error: any) => {
        console.log("Error Updating new item: " + error);
      });
  };
  const DeleteItemList = (Id: any) => {
    // Define the URL of the SharePoint site and the list name
    const webUrl = ContextService.GetUrl();
    const listName = "RohitFirstListTesting";

    // Define the new item data as an object

    const requestOptions: ISPHttpClientOptions = {
      headers: {
        Accept: "application/json;odata=nometadata",
        "Content-type": "application/json;odata=nometadata",
        "odata-version": "",
        "IF-MATCH": "*",
        "X-Http-Method": "DELETE",
      },
    };
    // Define the endpoint URL for adding a new item to the list
    const endpointUrl = `${webUrl}/_api/web/lists/getbytitle('${listName}')/items(${Id})`;

    // Send the POST request to create the new item
    ContextService.GetSPContext()
      .post(endpointUrl, SPHttpClient.configurations.v1, requestOptions)
      .then((response: SPHttpClientResponse) => {
        if (response.ok) {
          console.log("Deleted successfully");
        } else {
          console.log("Error Deleting item: " + response.statusText);
        }
      })
      .catch((error: any) => {
        console.log("Error Deleteing item: " + error);
      });
  };
  const GetDataTable = (e: any) => {
    let userdata = document.querySelectorAll(`.${e.target.className}`);
    let Uname = userdata[0].innerHTML;
    let Pass = userdata[1].innerHTML;
    let Id = userdata[1].id;

    UpdateChange(Id, Uname, Pass);
  };
  const DeleteUser = (e: any) => {
    confirm("Are You Want Delete This user") && DeleteItemList(e.target.id);
  };
  return (
    <div style={{ width: "100%" }}>
      <form onChange={handelChange} className="form">
        <input name="username" type="text" value={data.username} />
        <input name="pwd" type="password" value={data.pwd} />
        <button onClick={handleClick}>submit</button>
      </form>
      y6
      <table style={{ width: "100%" }}>
        <tr>
          <td style={{ fontWeight: 800, color: "red" }}>ID</td>
          <td style={{ fontWeight: 800, color: "red" }}>UserName</td>
          <td style={{ fontWeight: 800, color: "red" }}>Password</td>
          <td style={{ fontWeight: 800, color: "red" }}>UserAction</td>
        </tr>
        {listData.map((item: any, index: any) => {
          return (
            <tr>
              <td className="tableDataID">{index + 1}</td>
              <td
                // ref={fieldRef}
                className={`class` + item.ID}
                contentEditable="true"
                onBlur={GetDataTable}
              >
                {item.UserName}
              </td>
              <td
                className={`class` + item.ID}
                contentEditable="true"
                onBlur={GetDataTable}
                id={item.ID}
              >
                {item.Password}
              </td>
              <td>
                <button id={item.ID} onClick={DeleteUser}>
                  Delete
                </button>
              </td>
            </tr>
          );
        })}
      </table>
      <CheckTheme />
    </div>
  );
};
export default CurdOps;
