import * as React from "react";
import styles from "./Calendar.module.scss";
import { ICalendarProps } from "./ICalendarProps";
import { ICalendarState } from "./ICalendarState";
import { escape } from "@microsoft/sp-lodash-subset";
import * as moment from "moment-timezone";
import * as strings from "CalendarWebPartStrings";
import "react-big-calendar/lib/css/react-big-calendar.css";
import { MSGraphClient } from "@microsoft/sp-http";
require("./calendar.css");
import {
  CommunicationColors,
  FluentCustomizations,
  FluentTheme,
} from "@uifabric/fluent-theme";
import Year from "./Year";

import { Calendar as MyCalendar, momentLocalizer } from "react-big-calendar";

import {
  Customizer,
  IPersonaSharedProps,
  Persona,
  PersonaSize,
  PersonaPresence,
  HoverCard,
  IHoverCard,
  IPlainCardProps,
  HoverCardType,
  DefaultButton,
  DocumentCard,
  DocumentCardActivity,
  DocumentCardDetails,
  DocumentCardPreview,
  DocumentCardTitle,
  IDocumentCardPreviewProps,
  IDocumentCardPreviewImage,
  DocumentCardType,
  Label,
  ImageFit,
  IDocumentCardLogoProps,
  DocumentCardLogo,
  DocumentCardImage,
  Icon,
  Spinner,
  SpinnerSize,
  MessageBar,
  MessageBarType,
} from "office-ui-fabric-react";
import { EnvironmentType } from "@microsoft/sp-core-library";
import { mergeStyleSets } from "office-ui-fabric-react/lib/Styling";
import { WebPartTitle } from "@pnp/spfx-controls-react/lib/WebPartTitle";
import { Placeholder } from "@pnp/spfx-controls-react/lib/Placeholder";
import { WebPartContext } from "@microsoft/sp-webpart-base";
import { DisplayMode } from "@microsoft/sp-core-library";
import spservices from "../../../services/spservices";
import { stringIsNullOrEmpty } from "@pnp/common";
import { Event } from "../../../controls/Event/event";
import { IPanelModelEnum } from "../../../controls/Event/IPanelModeEnum";
import { IEventData } from "./../../../services/IEventData";
// import { IUserPermissions } from "./../../../services/IUserPermissions";
// import { Views } from "@pnp/pnpjs";
import { event } from "jquery";

//const localizer = BigCalendar.momentLocalizer(moment);
const localizer = momentLocalizer(moment);
/**
 * @export
 * @class Calendar
 * @extends {React.Component<ICalendarProps, ICalendarState>}
 */
export default class Calendar extends React.Component<
  ICalendarProps,
  ICalendarState
> {
  private spService: spservices = null;
  // private userListPermissions: IUserPermissions = undefined;
  public constructor(props) {
    super(props);

    this.state = {
      showDialog: false,
      eventData: [],
      selectedEvent: undefined,
      isloading: true,
      hasError: false,
      errorMessage: "",
      sAllEvents: [],
    };

    this.onDismissPanel = this.onDismissPanel.bind(this);
    this.onSelectEvent = this.onSelectEvent.bind(this);
    this.onSelectSlot = this.onSelectSlot.bind(this);
    this.spService = new spservices(this.props.context);
    moment.locale(
      this.props.context.pageContext.cultureInfo.currentUICultureName
    );
  }

  private onDocumentCardClick(ev: React.SyntheticEvent<HTMLElement, Event>) {
    ev.preventDefault();
    ev.stopPropagation();
  }
  /**
   * @private
   * @param {*} event
   * @memberof Calendar
   */
  private onSelectEvent(event: any) {
    // this.props.context.msGraphClientFactory
    //   .getClient()
    //   .then((client: MSGraphClient) => {
    //     client
    //       .api("/users")
    //       .filter(`mail eq '${event.attendeesEmail}'`)
    //       .select("id")
    //       .get((err, res) => {
    //         if (err) {
    //           console.error("Error while getting user by email:", err);
    //           return;
    //         }

    //         const user = res.value[0]; // assuming only one user with the given email
    //         const attendeesID = user.id;

    //         console.log("User ID:", attendeesID);
    //       });
    //   });

    this.setState({
      showDialog: true,
      selectedEvent: event,
      panelMode: IPanelModelEnum.edit,
    });

    console.log("Selected Event:", event);
  }

  /**
   *
   * @private
   * @param {boolean} refresh
   * @memberof Calendar
   */
  private async onDismissPanel(refresh: boolean) {
    this.setState({ showDialog: false });
    if (refresh === true) {
      this.setState({ isloading: true });
      // await this.loadEvents();
      await this.loadOutlookEvents();
      this.setState({ isloading: false });
    }
  }
  /**
   * @private
   * @memberof Calendar
   */
  // private async loadEvents() {
  //   try {
  //     // Teste Properties
  //     if (
  //       !this.props.list ||
  //       !this.props.siteUrl ||
  //       !this.props.eventStartDate.value ||
  //       !this.props.eventEndDate.value
  //     )
  //       return;

  //     this.userListPermissions = await this.spService.getUserPermissions(
  //       this.props.siteUrl,
  //       this.props.list
  //     );
  //     const eventsData: IEventData[] = await this.spService.getEvents(
  //       escape(this.props.siteUrl),
  //       escape(this.props.list),
  //       this.props.eventStartDate.value,
  //       this.props.eventEndDate.value
  //     );
  //     this.setState({
  //       eventData: eventsData,
  //       hasError: false,
  //       errorMessage: "",
  //     });
  //     console.log(eventsData);
  //   } catch (error) {
  //     this.setState({
  //       hasError: true,
  //       errorMessage: error.message,
  //       isloading: false,
  //     });
  //   }
  // }

  public async loadOutlookEvents() {
    let lAllEventsData = [],
      lAllOptions = [],
      NewArray = [],
      array = [];
    let index = 0;
    const now = new Date();
    const sixMonthsAgo = new Date();
    const twoYearsFuture = new Date();
    twoYearsFuture.setFullYear(now.getFullYear() + 2);
    sixMonthsAgo.setMonth(now.getMonth() - 6);
    // console.log(
    //   "Three months ago: " +
    //     threeMonthsAgo +
    //     "two years future: " +
    //     twoYearsFuture
    // );

    // const filterstatement = `(recurrence eq null or type eq 'seriesMaster')`;

    const filterstatement = `start/dateTime ge '${sixMonthsAgo.toISOString()}' and start/dateTime le '${twoYearsFuture.toISOString()}'`;
    console.log(filterstatement);
    const filterDate = new Date().toISOString();
    this.props.context.msGraphClientFactory
      .getClient()
      .then((client: MSGraphClient) => {
        client
          .api("me/calendar/events")
          .filter(filterstatement)
          // .orderby("createdDateTime")
          .top(5000)
          // .select("subject,organizer,start,end")
          .get((err, res?: any) => {
            if (err) {
              console.log("Error in fetching events from outlook: ", err);
              return;
            }

            console.log("All Events as res: ", res);

            // this.spService.AddOutlookEventstoList(res);

            res.value.forEach(async (element, i) => {
              // console.log("Events: ", element);
              index = i;

              // Getting email address of attendess
              const attendeesEmail = [];
              for (let j = 0; j < element.attendees.length; j++) {
                attendeesEmail.push(element.attendees[j].emailAddress.address);
              }

              //getting id of attendess
              // try {
              //   // const attendeesID = await this.spService.getIdByUserEmail(
              //   //   attendeesEmail,
              //   //   this.props.siteUrl
              //   // );
              //   // console.log(attendeesID);
              //   // ::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
              //   // this.props.context.msGraphClientFactory
              //   //   .getClient()
              //   //   .then((client: MSGraphClient) => {
              //   //     client
              //   //       .api("/users")
              //   //       .filter(`mail eq '${attendeesEmail}'`)
              //   //       .select("id")
              //   //       .get((err, res) => {
              //   //         if (err) {
              //   //           console.error(err);
              //   //           return;
              //   //         }
              //   //         const attendeesID = res.value[0].id;
              //   //         console.log(
              //   //           `User ID for email ${attendeesEmail}: ${attendeesID}`
              //   //         );
              //   //       });
              //   //   });
              // } catch (err) {
              //   console.log(err);
              // }

              // console.log(element.subject);
              const timezone = "Asia/Kolkata";
              // console.log("sD before format", element.start.dateTime);
              let startDate = moment
                .utc(element.start.dateTime)
                .tz(timezone)
                .format("YYYY-MM-DD[T]HH:mm:ss");
              // console.log(startDate);

              let endDate = moment
                .utc(element.end.dateTime)
                .tz(timezone)
                .format("YYYY-MM-DD[T]HH:mm:ss");

              // for (const ArrayofAttendess of element.attendees) {
              //   ArrayofAttendess = element.attendees;
              // }

              if (element.type === "seriesMaster") {
                element.forEach((item) => {
                  console.log("Item >>>>>>>>>>>>>>>>>>>>>", item);
                });
                lAllEventsData.push({
                  // Id: element.id,
                  id: element.id,
                  // ID: element.id,
                  title: element.subject,
                  Description: element.bodyPreview,
                  location: element.location.displayName,
                  EventDate: new Date(startDate),
                  EndDate: new Date(endDate),
                  // color: "",
                  // ownerInitial: "",
                  // ownerInitial: element.organizer.emailAddress.name,
                  // ownerPhoto: "#",
                  // ownerEmail: element.organizer.emailAddress.address,
                  ownerName: element.organizer.emailAddress.name,
                  fAllDayEvent: element !== null ? element.isAllDay : "",
                  attendes: element.attendees,
                  attendeesID: element.attendeesID,
                  attendessEmail: attendeesEmail,
                  geolocation: element.locations.displayName,
                  // Category: element.categories,
                  Category: "Test",
                  // Duration: 500,

                  fRecurrence: element !== null ? element.recurrence : "",

                  // Rpattern:
                  //   element.recurrence !== undefined
                  //     ? element.recurrence.pattern
                  //     : "",
                  Type: element.type,
                  EventType: "1",
                  iCalUId: element.iCalUId,
                  UID: element.iCalUId,
                  // RecurrenceID: element.RecurrenceID
                  //   ? element.RecurrenceID
                  //   : undefined,
                  // MasterSeriesItemID: element.seriesMasterId,
                  recurrenceInterval:
                    element.recurrence !== null
                      ? element.recurrence.pattern.interval
                      : "",
                  recurrenceRangeNumber:
                    element.recurrence !== null
                      ? element.recurrence.range.numberOfOccurrences
                      : "",
                  recurrenceRangeType:
                    element.recurrence !== null
                      ? element.recurrence.range.type
                      : "",
                  recurrenceStartTime:
                    element.recurrence !== null
                      ? element.recurrence.range.startDate
                      : "",
                  recurrenceEndTime:
                    element.recurrence !== null
                      ? element.recurrence.range.endDate
                      : "",
                  recurrenceTimeZone:
                    element.recurrence !== null
                      ? element.recurrence.range.recurrenceTimeZone
                      : "",
                  recurrencePattern:
                    element.recurrence !== null
                      ? element.recurrence.pattern
                      : "",
                  recurrencePatternType:
                    element.recurrence !== null
                      ? element.recurrence.pattern.type
                      : "",
                });
                console.log("This event is SeriesMaster >>>>>>>>>>>>>>>>>>>>");
              } else {
                lAllEventsData.push({
                  // Id: element.id,
                  id: element.id,
                  // ID: element.id,
                  title: element.subject,
                  Description: element.bodyPreview,
                  location: element.location.displayName,
                  EventDate: new Date(startDate),
                  EndDate: new Date(endDate),
                  // color: "",
                  // ownerInitial: "",
                  // ownerInitial: element.organizer.emailAddress.name,
                  // ownerPhoto: "#",
                  // ownerEmail: element.organizer.emailAddress.address,
                  ownerName: element.organizer.emailAddress.name,
                  fAllDayEvent: element !== null ? element.isAllDay : "",
                  attendes: element.attendees,
                  attendeesID: element.attendeesID,
                  attendessEmail: attendeesEmail,
                  geolocation: element.locations.displayName,
                  // Category: element.categories,
                  Category: "Test",
                  // Duration: 500,

                  fRecurrence: element !== null ? element.recurrence : "",

                  // Rpattern:
                  //   element.recurrence !== undefined
                  //     ? element.recurrence.pattern
                  //     : "",
                  Type: element.type,
                  EventType: "1",
                  iCalUId: element.iCalUId,
                  UID: element.iCalUId,
                  // RecurrenceID: element.RecurrenceID
                  //   ? element.RecurrenceID
                  //   : undefined,
                  // MasterSeriesItemID: element.seriesMasterId,
                  recurrenceInterval:
                    element.recurrence !== null
                      ? element.recurrence.pattern.interval
                      : "",
                  recurrenceRangeNumber:
                    element.recurrence !== null
                      ? element.recurrence.range.numberOfOccurrences
                      : "",
                  recurrenceRangeType:
                    element.recurrence !== null
                      ? element.recurrence.range.type
                      : "",
                  recurrenceStartTime:
                    element.recurrence !== null
                      ? element.recurrence.range.startDate
                      : "",
                  recurrenceEndTime:
                    element.recurrence !== null
                      ? element.recurrence.range.endDate
                      : "",
                  recurrenceTimeZone:
                    element.recurrence !== null
                      ? element.recurrence.range.recurrenceTimeZone
                      : "",
                  recurrencePattern:
                    element.recurrence !== null
                      ? element.recurrence.pattern
                      : "",
                  recurrencePatternType:
                    element.recurrence !== null
                      ? element.recurrence.pattern.type
                      : "",
                });
              }
              // lAllEventsData.push({
              //   // Id: element.id,
              //   id: element.id,
              //   // ID: element.id,
              //   title: element.subject,
              //   Description: element.bodyPreview,
              //   location: element.location.displayName,
              //   EventDate: new Date(startDate),
              //   EndDate: new Date(endDate),
              //   // color: "",
              //   // ownerInitial: "",
              //   // ownerInitial: element.organizer.emailAddress.name,
              //   // ownerPhoto: "#",
              //   // ownerEmail: element.organizer.emailAddress.address,
              //   ownerName: element.organizer.emailAddress.name,
              //   fAllDayEvent: element !== null ? element.isAllDay : "",
              //   attendes: element.attendees,
              //   attendeesID: element.attendeesID,
              //   attendessEmail: attendeesEmail,
              //   geolocation: element.locations.displayName,
              //   // Category: element.categories,
              //   Category: "Test",
              //   // Duration: 500,

              //   fRecurrence: element !== null ? element.recurrence : "",

              //   // Rpattern:
              //   //   element.recurrence !== undefined
              //   //     ? element.recurrence.pattern
              //   //     : "",
              //   Type: element.type,
              //   EventType: "1",
              //   iCalUId: element.iCalUId,
              //   UID: element.iCalUId,
              //   // RecurrenceID: element.RecurrenceID
              //   //   ? element.RecurrenceID
              //   //   : undefined,
              //   // MasterSeriesItemID: element.seriesMasterId,
              //   recurrenceInterval:
              //     element.recurrence !== null
              //       ? element.recurrence.pattern.interval
              //       : "",
              //   recurrenceRangeNumber:
              //     element.recurrence !== null
              //       ? element.recurrence.range.numberOfOccurrences
              //       : "",
              //   recurrenceRangeType:
              //     element.recurrence !== null
              //       ? element.recurrence.range.type
              //       : "",
              //   recurrenceStartTime:
              //     element.recurrence !== null
              //       ? element.recurrence.range.startDate
              //       : "",
              //   recurrenceEndTime:
              //     element.recurrence !== null
              //       ? element.recurrence.range.endDate
              //       : "",
              //   recurrenceTimeZone:
              //     element.recurrence !== null
              //       ? element.recurrence.range.recurrenceTimeZone
              //       : "",
              //   recurrencePattern:
              //     element.recurrence !== null ? element.recurrence.pattern : "",
              //   recurrencePatternType:
              //     element.recurrence !== null
              //       ? element.recurrence.pattern.type
              //       : "",
              // });
            });

            this.setState({ eventData: lAllEventsData });

            console.log("eventData:", this.state.eventData);
          });
      });
  }

  /**
   * @memberof Calendar
   */
  public async componentDidMount() {
    this.setState({ isloading: true });
    // await this.loadEvents();
    await this.loadOutlookEvents();
    this.setState({ isloading: false });
  }

  /**
   *
   * @param {*} error
   * @param {*} errorInfo
   * @memberof Calendar
   */
  public componentDidCatch(error: any, errorInfo: any) {
    this.setState({ hasError: true, errorMessage: errorInfo.componentStack });
  }
  /**
   *
   *
   * @param {ICalendarProps} prevProps
   * @param {ICalendarState} prevState
   * @memberof Calendar
   */
  public async componentDidUpdate(
    prevProps: ICalendarProps,
    prevState: ICalendarState
  ) {
    if (
      !this.props.list ||
      !this.props.siteUrl ||
      !this.props.eventStartDate.value ||
      !this.props.eventEndDate.value
    )
      return;
    // Get  Properties change
    if (
      prevProps.list !== this.props.list ||
      this.props.eventStartDate.value !== prevProps.eventStartDate.value ||
      this.props.eventEndDate.value !== prevProps.eventEndDate.value
    ) {
      this.setState({ isloading: true });
      // await this.loadEvents();
      await this.loadOutlookEvents();
      this.setState({ isloading: false });
    }
  }
  /**
   * @private
   * @param {*} { event }
   * @returns
   * @memberof Calendar
   */
  private renderEvent({ event }) {
    const previewEventIcon: IDocumentCardPreviewProps = {
      previewImages: [
        {
          // previewImageSrc: event.ownerPhoto,
          previewIconProps: {
            iconName:
              event.fRecurrence !== undefined ? "RecurringEvent" : "Calendar",
            styles: { root: { color: event.color } },
            className: styles.previewEventIcon,
          },
          height: 43,
        },
      ],
    };
    const EventInfo: IPersonaSharedProps = {
      imageInitials: event.ownerInitial,
      imageUrl: event.ownerPhoto,
      text: event.title,
    };

    /**
     * @returns {JSX.Element}
     */
    const onRenderPlainCard = (): JSX.Element => {
      return (
        <div className={styles.plainCard}>
          <DocumentCard className={styles.Documentcard}>
            <div>
              <DocumentCardPreview {...previewEventIcon} />
            </div>
            <DocumentCardDetails>
              <div className={styles.DocumentCardDetails}>
                <DocumentCardTitle
                  title={event.title}
                  shouldTruncate={true}
                  className={styles.DocumentCardTitle}
                  styles={{ root: { color: event.color } }}
                />
              </div>
              {moment(event.EventDate).format("YYYY/MM/DD") !==
              moment(event.EndDate).format("YYYY/MM/DD") ? (
                <span className={styles.DocumentCardTitleTime}>
                  {moment(event.EventDate).format("dddd")} -{" "}
                  {moment(event.EndDate).format("dddd")}{" "}
                </span>
              ) : (
                <span className={styles.DocumentCardTitleTime}>
                  {moment(event.EventDate).format("dddd")}{" "}
                </span>
              )}
              <span className={styles.DocumentCardTitleTime}>
                {moment(event.EventDate).format("HH:mm")}H -{" "}
                {moment(event.EndDate).format("HH:mm")}H
              </span>
              <Icon
                iconName="MapPin"
                className={styles.locationIcon}
                style={{ color: event.color }}
              />
              <DocumentCardTitle
                title={`${event.location}`}
                shouldTruncate={true}
                showAsSecondaryTitle={true}
                className={styles.location}
              />
              <div style={{ marginTop: 20 }}>
                <DocumentCardActivity
                  activity={strings.EventOwnerLabel}
                  people={[
                    {
                      name: event.ownerName,
                      profileImageSrc: event.ownerPhoto,
                      initialsColor: event.color,
                    },
                  ]}
                />
              </div>
            </DocumentCardDetails>
          </DocumentCard>
        </div>
      );
    };

    return (
      <div style={{ height: 22 }}>
        <HoverCard
          cardDismissDelay={1000}
          type={HoverCardType.plain}
          plainCardProps={{ onRenderPlainCard: onRenderPlainCard }}
          onCardHide={(): void => {}}
        >
          <Persona
            {...EventInfo}
            size={PersonaSize.size24}
            presence={PersonaPresence.none}
            coinSize={22}
            initialsColor={event.color}
          />
        </HoverCard>
      </div>
    );
  }
  /**
   *
   *
   * @private
   * @memberof Calendar
   */
  private onConfigure() {
    // Context of the web part
    this.props.context.propertyPane.open();
  }

  /**
   * @param {*} { start, end }
   * @memberof Calendar
   */
  public async onSelectSlot({ start, end }) {
    // if (!this.userListPermissions.hasPermissionAdd) return;
    this.setState({
      showDialog: true,
      startDateSlot: start,
      endDateSlot: end,
      selectedEvent: undefined,
      panelMode: IPanelModelEnum.add,
    });
  }

  /**
   *
   * @param {*} event
   * @param {*} start
   * @param {*} end
   * @param {*} isSelected
   * @returns {*}
   * @memberof Calendar
   */
  public eventStyleGetter(event, start, end, isSelected): any {
    let style: any = {
      backgroundColor: "white",
      borderRadius: "0px",
      opacity: 1,
      color: event.color,
      borderWidth: "1.1px",
      borderStyle: "solid",
      borderColor: event.color,
      borderLeftWidth: "6px",
      display: "block",
    };

    return {
      style: style,
    };
  }

  /**
   *
   * @param {*} date
   * @memberof Calendar
   */
  public dayPropGetter(date: Date) {
    return {
      className: styles.dayPropGetter,
    };
  }

  /**
   *
   * @returns {React.ReactElement<ICalendarProps>}
   * @memberof Calendar
   */
  public render(): React.ReactElement<ICalendarProps> {
    return (
      <Customizer {...FluentCustomizations}>
        <div
          className={styles.calendar}
          style={{ backgroundColor: "white", padding: "20px" }}
        >
          <WebPartTitle
            displayMode={this.props.displayMode}
            title={this.props.title}
            updateProperty={this.props.updateProperty}
          />
          {!this.props.list ||
          !this.props.eventStartDate.value ||
          !this.props.eventEndDate.value ? (
            <Placeholder
              iconName="Edit"
              iconText={strings.WebpartConfigIconText}
              description={strings.WebpartConfigDescription}
              buttonLabel={strings.WebPartConfigButtonLabel}
              hideButton={this.props.displayMode === DisplayMode.Read}
              onConfigure={this.onConfigure.bind(this)}
            />
          ) : // test if has errors
          this.state.hasError ? (
            <MessageBar messageBarType={MessageBarType.error}>
              {this.state.errorMessage}
            </MessageBar>
          ) : (
            // show Calendar
            // Test if is loading Events
            <div>
              {this.state.isloading ? (
                <Spinner
                  size={SpinnerSize.large}
                  label={strings.LoadingEventsLabel}
                />
              ) : (
                <div className={styles.container}>
                  <MyCalendar
                    dayPropGetter={this.dayPropGetter}
                    localizer={localizer}
                    selectable
                    events={this.state.eventData}
                    startAccessor="EventDate"
                    endAccessor="EndDate"
                    eventPropGetter={this.eventStyleGetter}
                    onSelectSlot={this.onSelectSlot}
                    components={{
                      event: this.renderEvent,
                    }}
                    onSelectEvent={this.onSelectEvent}
                    defaultDate={moment().startOf("day").toDate()}
                    views={{
                      day: true,
                      week: true,
                      month: true,
                      agenda: true,
                      work_week: Year,
                    }}
                    messages={{
                      today: strings.todayLabel,
                      previous: strings.previousLabel,
                      next: strings.nextLabel,
                      month: strings.monthLabel,
                      week: strings.weekLabel,
                      day: strings.dayLable,
                      showMore: (total) => `+${total} ${strings.showMore}`,
                      work_week: strings.yearHeaderLabel,
                    }}
                    // onView={(views) => {
                    //   console.log(`View changed to: ${views}`);
                    // }}
                  />
                </div>
              )}
            </div>
          )}

          {this.state.showDialog && (
            <Event
              event={this.state.selectedEvent}
              panelMode={this.state.panelMode}
              onDissmissPanel={this.onDismissPanel}
              showPanel={this.state.showDialog}
              startDate={this.state.startDateSlot}
              endDate={this.state.endDateSlot}
              context={this.props.context}
              siteUrl={this.props.siteUrl}
              listId={this.props.list}
            />
          )}
        </div>
      </Customizer>
    );
  }
}
