import { WebPartContext } from "@microsoft/sp-webpart-base";

// IEmpApiProps.ts
export interface IEmpApiProps {
  description: string;
  isDarkTheme: boolean;
  environmentMessage: string;
  hasTeamsContext: boolean;
  userDisplayName: string;
  context: WebPartContext; // Pass the SharePoint context to the component
}
