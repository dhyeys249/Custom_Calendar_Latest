export interface IEventData {
  id?: number;
  Id?: number;
  ID?: number;
  title: string;
  Description?: any;
  location?: string;
  EventDate: Date;
  EndDate: Date;
  color?: string;
  ownerInitial?: string;
  ownerPhoto?: string;
  ownerEmail?: string;
  ownerName?: string;
  fAllDayEvent?: boolean;
  attendes?: any[];
  attendesID?: number[];
  attendesEmail?: string[];
  geolocation?: { Longitude: number; Latitude: number };
  Category?: string;
  Duration?: number;
  RecurrenceData?: string;
  fRecurrence?: string | boolean;
  Type?: string;
  EventType?: string;
  iCalUId?: string;
  UID?: string;
  RecurrenceID?: Date;
  MasterSeriesItemID?: string;
  Rpattern?: [];
  numberOfOccurrences?: number;
  recurrenceInterval?: number;
  recurrenceRangeNumber?: number;
  recurrenceRangeType?: number;
  recurrenceStartTime?: Date;
  recurrenceEndTime?: Date;
  recurrenceTimeZone?: string;
  recurrencePattern?: string;
  recurrencePatternType?: string;
}