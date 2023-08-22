export const getUsersList = async () => {
  const apiUrl = `https://sxmyf.sharepoint.com/sites/ArkitektzTest/_api/web/lists('e3dd271b-6d15-4d07-9c5d-520cb4c4fff4')/items`;

  const response = await fetch(
    // "https://sxmyf.sharepoint.com/sites/ArkitektzTest/lists/e3dd271b-6d15-4d07-9c5d-520cb4c4fff4"
    apiUrl,
    {
      method: "GET",
      headers: { Accept: "application/json;odata=nometadata" },
    }
  );

  const responseData = await response.json();
  return responseData.value;
};
