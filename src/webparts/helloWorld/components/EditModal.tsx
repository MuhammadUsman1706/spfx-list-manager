import * as React from "react";
import UseDataInputType from "../hooks/useDataInputTypes";
import { useContext } from "react";
import { useId } from "@fluentui/react-hooks";
import { getTheme, mergeStyleSets, FontWeights, Modal } from "@fluentui/react";
import { Store } from "../context/user-context";

import styles from "./UserForm.module.scss";

const theme = getTheme();

const contentStyles = mergeStyleSets({
  container: {
    display: "flex",
    flexFlow: "column nowrap",
    alignItems: "stretch",
  },
  header: [
    theme.fonts.xLargePlus,
    {
      flex: "1 1 auto",
      borderTop: `4px solid ${theme.palette.themePrimary}`,
      color: theme.palette.neutralPrimary,
      display: "flex",
      alignItems: "center",
      fontWeight: FontWeights.semibold,
      padding: "12px 12px 14px 24px",
    },
  ],
  heading: {
    color: theme.palette.neutralPrimary,
    fontWeight: FontWeights.semibold,
    fontSize: "inherit",
    margin: "0",
  },
  body: {
    flex: "4 4 auto",
    padding: "0 24px 24px 24px",
    overflowY: "hidden",
    selectors: {
      p: { margin: "14px 0" },
      "p:first-child": { marginTop: 0 },
      "p:last-child": { marginBottom: 0 },
    },
  },
});

const EditModal: React.FC<any> = ({ request, requestSP }): any => {
  const {
    state: { columns, selectedList, selectedSite, editItem },
    dispatch,
  } = useContext(Store);

  const hideModal = () => {
    dispatch({ type: "SET_EDIT_ITEM", payload: null });
  };

  console.log(editItem);

  const titleId = useId("title");

  const [dataTypesInputsHandler, formData]: any = UseDataInputType({
    _searchWithSP: requestSP,
    _searchWithAad: request,
  });

  const handleSubmit = async (event: React.FormEvent): Promise<any> => {
    event.preventDefault();
    console.log(formData);
    await request(
      `https://graph.microsoft.com/v1.0/sites/${selectedSite.id}/lists/${selectedList.id}/items/${editItem.id}`,
      null,
      null,
      "post",
      {
        body: JSON.stringify({ fields: formData }),
        headers: {
          "Content-Type": "application/json",
          "X-HTTP-Method": "PATCH",
          "If-Match": editItem.eTag,
        },
      }
    ).then(() => {
      request(
        `https://graph.microsoft.com/v1.0/sites/${selectedSite.id}/lists/${selectedList.id}/items?$expand=fields`,
        "SET_TABLE_LIST"
      );

      hideModal();
    });
  };

  return (
    <div>
      <Modal
        titleAriaId={titleId}
        isOpen={Boolean(editItem)}
        onDismiss={hideModal}
        isBlocking={false}
        containerClassName={contentStyles.container}
        // dragOptions={isDraggable ? dragOptions : undefined}
      >
        <form className={styles["user-form"]} onSubmit={handleSubmit}>
          {columns.map((field: any) =>
            dataTypesInputsHandler[field["@odata.type"]](field)
          )}
          <button type="submit">Submit</button>
        </form>
      </Modal>
    </div>
  );
};

export default EditModal;
