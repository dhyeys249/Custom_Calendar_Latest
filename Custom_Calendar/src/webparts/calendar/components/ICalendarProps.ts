import { DisplayMode } from "@microsoft/sp-core-library";
import { WebPartContext } from "@microsoft/sp-webpart-base";
import { IDateTimeFieldValue } from "@pnp/spfx-property-controls/lib/PropertyFieldDateTimePicker";
export interface ICalendarProps {
  pTitle: string;
  siteUrl: string;
  list: string;
  pDisplayMode: DisplayMode;
  eventStartDate: IDateTimeFieldValue;
  eventEndDate: IDateTimeFieldValue;
  pSpfxContext: WebPartContext;
  userDisplayName: string;
  pUpdateProperty: (value: string) => void;
  pContext: WebPartContext;
  context: any;
  checkPermission: any;
  uploadImage: any;
  backToHome: any;
  SiteUrl: any;
}
