import * as React from "react";
import { useContext } from "react";
import styles from "./TableList.module.scss"; // Import your SCSS file
import { Store } from "../context/user-context";
import EditModal from "./EditModal";

interface TableListProps {
  request: Function;
  selectedSite: any;
  selectedList: any;
}

interface property {
  Url: string;
  Description: string;
}

interface column {
  Title: string;
  StaticName: string;
}

interface field {
  Role: string;
  Title: string;
  Property: property;
  [key: string]: string | any;
}

interface row {
  fields: field;
  id: string;
}

const TableList: React.FC<TableListProps> = ({
  request,
  selectedList,
  selectedSite,
}) => {
  const {
    // state: { columns, tableList, selectedList, selectedSite, editItem },
    state: { columns, tableList },
    dispatch,
  } = useContext(Store);

  if (!tableList.length || !columns.length)
    return <h1 className={styles["no-data"]}>No data to display</h1>;

  console.log(columns, tableList);

  const deleteItemhandler = (event: any) => {
    const link = `https://graph.microsoft.com/v1.0/sites/${selectedSite.id}/lists/${selectedList.id}/items/${event?.target?.id}`;

    request(link, null, null, "post", {
      method: "DELETE",
      headers: { "X-HTTP-Method": "DELETE" },
    }).then((data: any) => {
      console.log(data);
      request(
        `https://graph.microsoft.com/v1.0/sites/${selectedSite.id}/lists/${selectedList.id}/items?$expand=fields`,
        "SET_TABLE_LIST"
      );
    });
  };

  const setEditItemHandler = (
    event: React.MouseEvent<HTMLButtonElement, MouseEvent>,
    row: row
  ) => {
    // const { id } = event.target as HTMLElement;

    dispatch({ type: "SET_EDIT_ITEM", payload: row });
  };

  return (
    <div className={styles["table-container"]}>
      <table className={styles["custom-table"]}>
        <thead>
          <tr>
            {columns.map((column, index) => (
              <th key={index}>{column.Title}</th>
            ))}
            <tr />
          </tr>
        </thead>
        <tbody>
          {tableList.map((row: row, rowIndex: number) => (
            <tr key={rowIndex}>
              {columns.map((column: column, cellIndex) => (
                <td key={cellIndex}>
                  {typeof row.fields[column.StaticName] === "object"
                    ? row?.fields[column.StaticName]?.Url
                    : row?.fields[column.StaticName]?.toString()}
                </td>
              ))}
              <td>
                <button id={row.id} onClick={deleteItemhandler}>
                  Delete
                </button>
              </td>
              <td>
                <button
                  id={row.id}
                  onClick={(event) => setEditItemHandler(event, row)}
                >
                  Edit
                </button>
              </td>
            </tr>
          ))}
        </tbody>
      </table>
    </div>
  );
};

export default TableList;
