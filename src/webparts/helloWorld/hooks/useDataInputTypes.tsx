import * as React from "react";
import { useContext, useMemo, useState } from "react";
import { TextField } from "@fluentui/react/lib/TextField";
import {
  ChoiceGroup,
  IChoiceGroupOption,
} from "@fluentui/react/lib/ChoiceGroup";
import { Dropdown, IDropdownOption } from "@fluentui/react/lib/Dropdown";
import { Checkbox } from "@fluentui/react";
import { Store } from "../context/user-context";

interface dataTypesInputsHandler {
  [key: string]: Function;
}

const UseDataInputType: React.FC<any> = ({
  _searchWithSP,
  _searchWithAad,
}): any => {
  const {
    state: { selectedSite, lookupsItems, userList, editItem },
  } = useContext(Store);

  const [formData, setFormData] = useState<any>({});

  const handleInputChange = (
    event: any,
    value: undefined | readonly string[] | number | string | boolean | any
  ): void => {
    console.log(value);

    const { name, id } = event.target;
    setFormData((prevData: any) => ({
      ...prevData,
      [name || id]: value,
    }));
  };

  const handlePersonChange = async (
    event: React.FormEvent<HTMLDivElement>,
    item: any
  ) => {
    const link = `${selectedSite?.webUrl}/_api/web/ensureUser`;

    const response = await _searchWithSP(link, null, null, "post", {
      body: JSON.stringify({ logonName: item.userPrincipalName }),
      headers: { "Content-Type": "application/json" },
    });

    handleInputChange(event, response.Id);
  };

  const dataTypesInputsHandler: dataTypesInputsHandler = useMemo(() => {
    return {
      "#SP.FieldText": (field: any) => (
        <TextField
          label={field.Title}
          type="text"
          name={field.StaticName}
          //   value={formData[field.StaticName]}
          onChange={handleInputChange}
          required={field.Required}
          defaultValue={editItem?.fields[field.StaticName]}
        />
      ),
      "#SP.FieldMultiLineText": (field: any) => (
        <TextField
          label={field.Title}
          type="text"
          name={field.StaticName}
          //   value={formData[field.StaticName]}
          onChange={handleInputChange}
          multiline
          required={field.Required}
          defaultValue={editItem?.fields[field.StaticName]}
        />
      ),
      "#SP.FieldNumber": (field: any) => (
        <TextField
          label={field.Title}
          type="number"
          name={field.StaticName}
          //   value={formData[field.StaticName]}
          onChange={handleInputChange}
          required={field.Required}
          defaultValue={editItem?.fields[field.StaticName]}
        />
      ),
      "#SP.FieldDateTime": (field: any) => (
        <TextField
          label={field.Title}
          type="date"
          name={field.StaticName}
          //   value={formData[field.StaticName]}
          onChange={handleInputChange}
          required={field.Required}
          defaultValue={editItem?.fields[field.StaticName]}
        />
      ),
      "#SP.FieldChoice": (field: any) => (
        <ChoiceGroup
          options={field.Choices.map((option: string, index: number) => ({
            key: index.toString(),
            text: option,
          }))}
          onChange={(event: any, option) =>
            handleInputChange(event, option.text)
          }
          label={field.Title}
          required={field.Required}
          name={field.Title}
          defaultValue={editItem?.fields[field.StaticName]}
        />
      ),
      "#SP.Field": (field: any) => (
        <Checkbox
          style={{ marginTop: "1rem" }}
          name={field.Title}
          label={field.Title}
          onChange={handleInputChange}
          defaultValue={editItem?.fields[field.StaticName]}
        />
      ),
      "#SP.FieldUrl": (field: any) => (
        <TextField
          label={field.Title}
          type="url"
          name={field.StaticName}
          //   value={formData[field.StaticName]}
          onChange={(event, value) =>
            handleInputChange(event, { Url: value, Description: "Google" })
          }
          required={field.Required}
          //   defaultValue={editItem?.fields[field.StaticName]}
        />
      ),
      "#SP.FieldLookup": (field: any) => {
        let value;
        const link = `${field["@odata.id"].split("Lists")[0]}/lists('${
          field.LookupList
        }')/items`;
        if (!lookupsItems[field?.LookupList]) {
          _searchWithSP(link, "SET_LOOKUPS_ITEMS", (value: any) => ({
            id: field?.LookupList,
            items: value,
          }))
            .then((data: any) => {
              value = data;
            })
            .catch((err: any) => console.log(err));
        } else {
          value = lookupsItems[field?.LookupList];
        }

        return (
          <Dropdown
            placeholder="Select a Lookup"
            label={field.Title}
            ariaLabel="Lookup Selection"
            onRenderOption={(option: any) => option[field?.LookupField]}
            options={value}
            id={`${field.StaticName}LookupId`}
            onChange={(event, item: any) => handleInputChange(event, item.Id)}
            selectedKey={formData[`${field.StaticName}LookupId`] || undefined}
            defaultValue={editItem?.fields[field.StaticName]}
          />
        );
      },
      "#SP.FieldUser": (field: any) => {
        let users;

        if (!userList?.length) {
          _searchWithAad(
            "https://graph.microsoft.com/v1.0/users",
            "SET_USERS_LIST"
          )
            .then((data: any) => (users = data))
            .catch((err: any) => console.log(err));
        } else {
          users = userList;
        }

        return (
          <Dropdown
            placeholder="Select a User"
            label={field.Title}
            ariaLabel="User Selection"
            onRenderOption={(option: any) => option?.displayName}
            options={users}
            id={`${field.StaticName}LookupId`}
            onChange={handlePersonChange}
            selectedKey={formData[`${field.StaticName}LookupId`] || undefined}
            defaultValue={editItem?.fields[field.StaticName]}
          />
        );
      },
    };
  }, [lookupsItems, userList, editItem]);

  return [dataTypesInputsHandler, formData];
};

export default UseDataInputType;
