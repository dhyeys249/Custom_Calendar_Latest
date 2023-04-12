import { WebPartContext } from "@microsoft/sp-webpart-base";
export interface IEventRecurrenceInfoProps {
  display: boolean;
  recurrenceData: any;
  startDate: Date;
  context: WebPartContext;
  siteUrl: string;
  returnRecurrenceData: (startDate: Date, recurrenceData: any) => void;
}
