import { WebPartContext } from "@microsoft/sp-webpart-base";
import { ClientMode } from "./ClientMode";

export interface IHelloWorldProps {
  description: string;
  isDarkTheme: boolean;
  environmentMessage: string;
  hasTeamsContext: boolean;
  userDisplayName: string;
  clientMode?: ClientMode;
  context?: WebPartContext;
  ctx: any;
  selectedList: any;
  selectedSite: any;
}
