import * as React from "react";
import UseDataInputType from "../hooks/useDataInputTypes";
import { useContext } from "react";
import { Store } from "../context/user-context";

import styles from "./UserForm.module.scss";

const UserForm: React.FC<any> = ({
  request,
  requestSP,
  selectedList,
  selectedSite,
}) => {
  const {
    // state: { columns, selectedList, selectedSite },
    state: { columns },
  } = useContext(Store);

  const [dataTypesInputsHandler, formData]: any = UseDataInputType({
    _searchWithSP: requestSP,
    _searchWithAad: request,
  });

  const handleSubmit = async (event: React.FormEvent): Promise<any> => {
    event.preventDefault();
    console.log(formData);
    await request(
      `https://graph.microsoft.com/v1.0/sites/${selectedSite.id}/lists/${selectedList.id}/items`,
      null,
      null,
      "post",
      {
        body: JSON.stringify({ fields: formData }),
        headers: { "Content-Type": "application/json" },
      }
    ).then(() =>
      request(
        `https://graph.microsoft.com/v1.0/sites/${selectedSite.id}/lists/${selectedList.id}/items?$expand=fields`,
        "SET_TABLE_LIST"
      )
    );

    // createUser(formData);
    // const apiUrl =
    //   "https://sxmyf.sharepoint.com/sites/ArkitektzTest/_api/web/lists('e3dd271b-6d15-4d07-9c5d-520cb4c4fff4')/items";
    // const body = JSON.stringify({ Title: formData.name, Role: formData.role });

    // if (formData.role && formData.name) {
    //   const spHttpClient: SPHttpClient = ctx.spHttpClient;
    //   await spHttpClient.post(apiUrl, SPHttpClient.configurations.v1, { body });
    //   refreshData();
    // }
  };

  if (!columns.length)
    return <h1 className={styles["no-data"]}>No data to submit</h1>;

  return (
    <form className={styles["user-form"]} onSubmit={handleSubmit}>
      {columns.map((field: any) =>
        dataTypesInputsHandler[field["@odata.type"]](field)
      )}
      <button type="submit">Submit</button>
    </form>
  );
};

export default UserForm;
