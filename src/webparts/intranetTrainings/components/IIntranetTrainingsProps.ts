import { WebPartContext } from "@microsoft/sp-webpart-base";

export interface IIntranetTrainingsProps {
  description: string;
  isDarkTheme: boolean;
  environmentMessage: string;
  hasTeamsContext: boolean;
  userDisplayName: string;
  listTitle: string;
  context: WebPartContext;
  webPartTitle: string;
  cardBackgroundColor: string;
  cardBorderColor: string;
  titleColor: string;
  dateColor: string;
  useFullWidth: boolean;
  itemsPerRow: number;
  showScheduledDate: boolean;
  titleFontColor: string;
  titleFontWeight: string;
  titleFontSize: number;
  titleBackgroundColor: string;
  cardHeight: number;
   // Category properties
  showCategory: boolean;
  categoryColor: string;
  categoryBgColor: string;
  maxRowsBeforeScroll: number; 
  dateBackgroundColor: string;
  enableScroll: boolean; 
  
}

export interface ITrainingItems {
  Id: number;
  Title: string;
  Link: string | { Url: string };
  Status: boolean;
AttachmentFiles?: { ServerRelativeUrl: string; FileName: string }[]; 
Created: string;
ScheduledDate?: string;
 Category?: string;
}

