import { WebPartContext } from "@microsoft/sp-webpart-base";

export interface IHrPandalogusaProps {
  description: string;
  docLibName:string;
  commentsListName:string;
  isDarkTheme: boolean;
  environmentMessage: string;
  hasTeamsContext: boolean;
  userDisplayName: string;
  context:WebPartContext;
}
