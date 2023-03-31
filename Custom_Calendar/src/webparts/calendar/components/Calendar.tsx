import * as React from "react";
import styles from "./Calendar.module.scss";
import { ICalendarProps } from "./ICalendarProps";
import { ICalendarState } from "./ICalendarState";
import { escape } from "@microsoft/sp-lodash-subset";
import * as moment from "moment-timezone";
import * as strings from "CalendarWebPartStrings";
import "react-big-calendar/lib/css/react-big-calendar.css";
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
import { IUserPermissions } from "./../../../services/IUserPermissions";
import { MSGraphClient } from "@microsoft/sp-http";
import * as MicrosoftGraph from "@microsoft/microsoft-graph-types";
import { graph } from "@pnp/graph";
import { Views, View } from "@pnp/sp";

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
  // handleShowMore = (events, date, view) => {
  //   if (view === "day") {
  //     console.log(`Showing more events for date ${date} in the day view`);
  //     // Do something with the day view
  //   } else {
  //     console.log(`Showing more events in the ${view} view`);
  //     // Handle other views
  //   }
  // };
  private spService: spservices = null;
  private userListPermissions: IUserPermissions = undefined;
  public constructor(props) {
    super(props);

    this.state = {
      sShowDialog: false,
      sEventData: [],
      sSelectedEvent: undefined,
      sIsloading: true,
      sHasError: false,
      sErrorMessage: "",
      sAllEvents: [],
      sSingleValueDropdown: "",
      sDropdownOptions: [],
      sIsDropdownSelected: false,

      contextitem: "",
      createmode: false,
      isUserAdmin: false,
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
    this.setState({
      sShowDialog: true,
      sSelectedEvent: event,
      sPanelMode: IPanelModelEnum.edit,
    });
  }

  /**
   *
   * @private
   * @param {boolean} refresh
   * @memberof Calendar
   */
  private async onDismissPanel(refresh: boolean) {
    this.setState({ sShowDialog: false });
    if (refresh === true) {
      this.setState({ sIsloading: true });
      // await this.loadEvents();

      await this.loadOutlookEvents();
      this.setState({ sIsloading: false });
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
  //       sEventData: eventsData,
  //       sHasError: false,
  //       sErrorMessage: "",
  //     });
  //   } catch (error) {
  //     this.setState({
  //       sHasError: true,
  //       sErrorMessage: error.message,
  //       sIsloading: false,
  //     });
  //   }
  // }
  // ::::::::::
  public async loadOutlookEvents() {
    let lAllEventsData = [],
      lAllOptions = [],
      NewArray = [],
      array = [];
    let index = 0;
    const now = new Date();
    const threeMonthsAgo = new Date();
    const twoYearsFuture = new Date();
    twoYearsFuture.setMonth(now.getMonth() + 24);
    threeMonthsAgo.setMonth(now.getMonth() - 3);
    console.log(
      "Three months ago: " +
        threeMonthsAgo +
        "two years future: " +
        twoYearsFuture
    );
    const filterDate = new Date().toISOString();
    this.props.context.msGraphClientFactory
      .getClient()
      .then((client: MSGraphClient) => {
        client
          .api("me/calendar/events")

          // .filter(
          //   `start/dateTime ge '${threeMonthsAgo.toISOString()}' and end/dateTime ge '${twoYearsFuture.toISOString()}'`
          // )
          .orderby("createdDateTime DESC")
          .top(5000)
          // .select("subject,organizer,start,end")
          .get((err, res?: any) => {
            if (err) {
              console.log("Error in fetching events from outlook: ", err);
              return;
            }

            console.log("All Events: ", res);

            // this.spService.AddOutlookEventstoList(res);

            res.value.forEach(async (element, i) => {
              // console.log("Events: ", element);
              index = i;
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

              lAllEventsData.push({
                id: element.id,
                title: element.subject,
                EventDate: startDate,
                start: startDate,
                end: endDate,
                EndDate: endDate,
                resource: element.attendees,
                isAllDay: element.isAllDay,
                attendees: element.attendees,
                categories: element.categories,
                recurrence: element.recurrence,
                type: element.type,
                iCalUId: element.iCalUId,
                ownerName: element.organizer.emailAddress.name,
                location: element.location.displayName,
              });
              // console.log("Attendees: ", element.attendees);
              // console.log("Categories: ", element.categories);
            });
            this.setState({ sAllEvents: lAllEventsData });
          });
      });

    // this.context.msGraphClientFactory
    //   .getClient()
    //   .then((client: MSGraphClient) => {
    //     client
    //       .api("me/events")
    //       .version("v1.0")
    //       .select("subject,organizer,start,end")
    //       .get((error, response: any, rawResponse?: any) => {
    //         if (error) {
    //           console.log(error);
    //           return;
    //         }
    //         const events: MicrosoftGraph.Event[] = response.value;
    //         console.log(events);
    //       });
    //   });
  }
  // :::::::::::::::::::
  /**
   * @memberof Calendar
   */
  public async componentDidMount() {
    this.setState({ sIsloading: true });
    // await this.loadEvents();
    await this.loadOutlookEvents();
    // await this.spService.AddOutlookEventstoList(this.context);
    this.setState({ sIsloading: false });
  }

  /**
   *
   * @param {*} error
   * @param {*} errorInfo
   * @memberof Calendar
   */
  public componentDidCatch(error: any, errorInfo: any) {
    this.setState({ sHasError: true, sErrorMessage: errorInfo.componentStack });
  }
  /**
   *
   *
   * @param {ICalendarProps} prevProps
   * @param {ICalendarState} prevState
   * @memberof Calendar
   */
  // public async componentDidUpdate(
  //   prevProps: ICalendarProps,
  //   prevState: ICalendarState
  // ) {
  //   if (
  //     !this.props.list ||
  //     !this.props.siteUrl ||
  //     !this.props.eventStartDate.value ||
  //     !this.props.eventEndDate.value
  //   )
  //     return;
  //   // Get  Properties change
  //   if (
  //     prevProps.list !== this.props.list ||
  //     this.props.eventStartDate.value !== prevProps.eventStartDate.value ||
  //     this.props.eventEndDate.value !== prevProps.eventEndDate.value
  //   ) {
  //     this.setState({ sIsloading: true });
  //     // await this.loadEvents();
  //     this.setState({ sIsloading: false });
  //   }
  // }
  /**
   * @private
   * @param { event }
   * @returns
   * @memberof Calendar
   */
  private renderEvent({ event }) {
    const previewEventIcon: IDocumentCardPreviewProps = {
      previewImages: [
        {
          // previewImageSrc: event.ownerPhoto,
          previewIconProps: {
            iconName: event.fRecurrence === "0" ? "Calendar" : "RecurringEvent",
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
      sShowDialog: true,
      sStartDateSlot: start,
      sEndDateSlot: end,
      sSelectedEvent: undefined,
      sPanelMode: IPanelModelEnum.add,
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
            displayMode={this.props.pDisplayMode}
            title={this.props.pTitle}
            updateProperty={this.props.pUpdateProperty}
          />
          {/* {!this.props.list ||
          !this.props.eventStartDate.value ||
          !this.props.eventEndDate.value ? (
            <Placeholder
              iconName="Edit"
              iconText={strings.WebpartConfigIconText}
              description={strings.WebpartConfigDescription}
              buttonLabel={strings.WebPartConfigButtonLabel}
              hideButton={this.props.pDisplayMode === DisplayMode.Read}
              onConfigure={this.onConfigure.bind(this)}
            />
          ) : // test if has errors
          this.state.sHasError ? (
            <MessageBar messageBarType={MessageBarType.error}>
              {this.state.sErrorMessage}
            </MessageBar>
          ) : (
            // show Calendar
            // Test if is loading Events */}
          <div>
            {/* {this.state.sIsloading ? (
                <Spinner
                  size={SpinnerSize.large}
                  label={strings.LoadingEventsLabel}
                />
              ) : ( */}
            <div className={styles.container}>
              {/* <div className={styles.calendarcontainer}> */}
              <MyCalendar
                dayPropGetter={this.dayPropGetter}
                localizer={localizer}
                selectable
                // events={this.state.eventData}
                events={this.state.sAllEvents}
                startAccessor="EventDate"
                endAccessor="EndDate"
                eventPropGetter={this.eventStyleGetter}
                onSelectSlot={this.onSelectSlot}
                // onShowMore={this.handleShowMore}
                components={{
                  event: this.renderEvent,
                }}
                // defaultView={"day"}
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

                // onShowMore={(events, date, view) => {
                //   if (view === "month") {
                //   }
                // }}
              />
            </div>
            {/* </div> */}
            {/* )} */}
          </div>
          {/* )} */}
          {this.state.sShowDialog && (
            <Event
              event={this.state.sSelectedEvent}
              panelMode={this.state.sPanelMode}
              onDissmissPanel={this.onDismissPanel}
              showPanel={this.state.sShowDialog}
              startDate={this.state.sStartDateSlot}
              endDate={this.state.sEndDateSlot}
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
