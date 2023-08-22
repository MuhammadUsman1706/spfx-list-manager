import * as React from "react";
import { createContext, useReducer, ReactNode, useEffect } from "react";
import { getUsersList } from "../api/index";

interface editItem {
  id: string;
  fields: any;
  eTag: string;
}

interface lookupsItems {
  [key: string]: any[];
}

interface StateType {
  sites: any[];
  lists: any[];
  selectedSite: any | null;
  selectedList: any;
  tableList: any[];
  columns: any[];
  lookupsItems: lookupsItems;
  userList: any[];
  editItem?: editItem;
  // Define your state properties here
}

interface ActionType {
  type: string;
  payload: any;
  // Define other action properties here
}

const initialState: StateType = {
  sites: [],
  lists: [],
  selectedSite: null,
  selectedList: null,
  tableList: [],
  columns: [],
  lookupsItems: {},
  userList: [],
  editItem: null,
};

export const Store = createContext<{
  state: StateType;
  dispatch: React.Dispatch<ActionType>;
}>({
  state: initialState,
  dispatch: () => {
    return null;
  },
});

const reducer = (state: StateType, action: ActionType): any => {
  switch (action.type) {
    case "SET_SITES":
      return { ...state, sites: action.payload };

    case "SET_LISTS":
      return { ...state, lists: action.payload };

    case "SET_TABLE_LIST":
      return { ...state, tableList: action.payload };

    case "SET_SELECTED_SITE":
      return { ...state, selectedSite: action.payload };

    case "SET_SELECTED_LIST":
      return { ...state, selectedList: action.payload };

    case "SET_COLUMNS":
      return { ...state, columns: action.payload };

    case "SET_LOOKUPS_ITEMS":
      return {
        ...state,
        lookupsItems: {
          ...state.lookupsItems,
          [action.payload.id]: action.payload.items,
        },
      };

    case "SET_USERS_LIST":
      return { ...state, userList: action.payload };

    case "SET_EDIT_ITEM":
      return { ...state, editItem: action.payload };

    default:
      return state;
  }
};

interface StoreProviderProps {
  children: ReactNode;
}

export const StoreProvider: React.FC<StoreProviderProps> = ({
  children,
}: StoreProviderProps) => {
  const [state, dispatch] = useReducer(reducer, initialState);

  // useEffect(() => {
  //   getUsersList()
  //     .then((data) => dispatch({ type: "USER_LIST", payload: data }))
  //     .catch((err) => console.log(err));
  // }, []);

  return (
    <Store.Provider value={{ state, dispatch }}>{children}</Store.Provider>
  );
};
