import { IPanelModelEnum } from "../../../controls/Event/IPanelModeEnum";
import { IEventData } from "./../../../services/IEventData";
export interface ICalendarState {
  sShowDialog: boolean;
  sEventData: IEventData[];
  sSelectedEvent: IEventData;
  sPanelMode?: IPanelModelEnum;
  sStartDateSlot?: Date;
  sEndDateSlot?: Date;
  sIsloading: boolean;
  sHasError: boolean;
  sErrorMessage: string;
  sAllEvents: any[];
  sDropdownOptions: any[];
  sSingleValueDropdown: string;
  sIsDropdownSelected: boolean;

  contextitem: string;
  createmode: boolean;
  isUserAdmin: boolean;
}
