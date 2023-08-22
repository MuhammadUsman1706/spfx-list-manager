import { IUserItem } from "./IUserItem";

export interface IHelloWorldState {
  users: Array<IUserItem>;
  searchFor: string;
}
