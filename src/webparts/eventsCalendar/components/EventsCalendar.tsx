import * as React from "react";
import { useEffect, useState } from "react";
import { MSGraphClientV3 } from "@microsoft/sp-http";
import * as MicrosoftGraph from "@microsoft/microsoft-graph-types";
import FullCalendar from "@fullcalendar/react";
import dayGridPlugin from "@fullcalendar/daygrid";
import timeGridPlugin from "@fullcalendar/timegrid";
import interactionPlugin from "@fullcalendar/interaction";
import { IEventsCalendarProps } from "./IEventsCalendarProps";
import styles from "./EventsCalendar.module.scss";
import OverlayTrigger from "react-bootstrap/OverlayTrigger";
import Popover from "react-bootstrap/Popover";

interface IFormattedEvent {
  subject: string;
  startDate: string;
  endDate: string;
  startTime: string;
  endTime: string;
  eventDate: string;
  bodyPreview?: string;
  joinUrl?: string;
}

var customStyles = `

    a {
      color: #000;
      text-decoration: none;

    }
    .fc .fc-button-primary:disabled {
      background-color: #7787A9;
      border-color: #7787A9;
      opacity: 1;
    }

    .fc .fc-button-primary {
      background-color: #293859;
      border-color: #293859;
    }

    .fc .fc-button-primary:not(:disabled).fc-button-active {
      back
    }

    :root {
      --fc-today-bg-color: #ececec;
      --fc-event-bg-color: #91afd9db;
    --fc-event-border-color: #91afd9db;
    }
  `;

const EventsCalendar: React.FC<IEventsCalendarProps> = (props: any) => {
  const [events, setEvents] = useState<MicrosoftGraph.Event[]>([]);

  useEffect(() => {
    props.context.msGraphClientFactory
      .getClient()
      .then((client: MSGraphClientV3) => {
        client
          .api("me/calendar/events")
          .version("v1.0")
          .select("*")
          .get((error: any, eventsResponse, rawResponse?: any) => {
            if (error) {
              console.error("Message is: " + error);
              return;
            }

            const calendarEvents: MicrosoftGraph.Event[] = eventsResponse.value;
            setEvents(
              calendarEvents.map((event) => ({
                ...event,
                joinUrl: event.onlineMeeting?.joinUrl || "",
                bodyPreview: event.bodyPreview || "", // Use bodyPreview if available, otherwise default to an empty string
              }))
            );

            console.log("CalendarEvents", calendarEvents);
          });
      });
  }, [props.context.msGraphClientFactory]);

  const eventContent = (eventInfo: any) => {
    const formattedEvent: IFormattedEvent = {
      subject: eventInfo.event.title,
      startDate: eventInfo.event.startStr,
      startTime: eventInfo.event.start.toLocaleTimeString(),
      endDate: eventInfo.event.endStr,
      endTime: eventInfo.event.end.toLocaleTimeString(),
      eventDate: eventInfo.event.start.toString(),
      bodyPreview: eventInfo.event.extendedProps.bodyPreview,
      joinUrl: eventInfo.event.extendedProps.joinUrl,
    };
    //console.log("Join URL:", formattedEvent.joinUrl);
    // console.log("Formatted Event", formattedEvent);
    // console.log("EventInfo", eventInfo);
    // console.log("Event Content", eventContent);

    const popover = (
      <Popover
        id={`popover-${formattedEvent.startDate}`}
        className={styles.popoverBox}
      >
        <Popover.Header as="h3" className={styles.popheader}>
          <b>Calendar</b> - <span>{props.context.pageContext.user.email}</span>
        </Popover.Header>
        <Popover.Body>
          <div className={styles.popBody}>
            <div className={styles.popheading}>
              <img
                src={require("../assets/Icon1.svg")}
                alt="Icon"
                className={styles.popoverIcon}
              />
              <span className={styles.contentStyle}>
                {formattedEvent.subject}
              </span>
            </div>
            <div className={styles.popContent}>
              <img
                src={require("../assets/Icon2.svg")}
                alt="Icon"
                className={styles.popoverIcon}
              />
              <span className={styles.contentStyle}>
                {`${formattedEvent.eventDate.substring(
                  0,
                  3
                )}, ${formattedEvent.eventDate.substring(
                  4,
                  10
                )} ${formattedEvent.startTime.substring(
                  0,
                  5
                )} - ${formattedEvent.endTime.substring(0, 5)}`}
              </span>
            </div>
            <div
              className={styles.popContent}
              style={{ display: formattedEvent.bodyPreview ? "flex" : "none" }}
            >
              <img
                src={require("../assets/Icon3.svg")}
                alt="Icon"
                className={styles.popoverIcon}
              />
              <p className={styles.contentStyle}>
                {formattedEvent.bodyPreview}
              </p>
            </div>
            <div style={{ display: formattedEvent.joinUrl ? "flex" : "none" }}>
              <button className={styles.joinBtn}>
                <a href={formattedEvent.joinUrl} target="_blank">
                  Join
                </a>
              </button>
            </div>
          </div>
        </Popover.Body>
      </Popover>
    );
    return (
      <OverlayTrigger
        trigger="click"
        placement="right"
        overlay={popover}
        rootClose={true}
      >
        <button className={styles.popoverButton}>
          <span>{formattedEvent.startTime.substring(0, 5)} </span>
          <b> {formattedEvent.subject}</b>
        </button>
      </OverlayTrigger>
    );
  };

  return (
    <div className={styles.calendarApp}>
      <style>{customStyles}</style>
      <div className={styles.calendarAppMain}>
        <FullCalendar
          plugins={[dayGridPlugin, timeGridPlugin, interactionPlugin]}
          headerToolbar={{
            left: "prev,next today",
            center: "title",
            right: "dayGridMonth,timeGridWeek,timeGridDay",
          }}
          initialView="dayGridMonth"
          customButtons={{
            customPrev: { text: "Prev" },
            customNext: { text: "Next" },
            customToday: { text: "Today" },
          }}
          buttonText={{
            prev: "<",
            next: ">",
            today: "Today",
            dayGridMonth: "Month",
            timeGridWeek: "Week",
            timeGridDay: "Day",
          }}
          events={events.map((event: any) => ({
            title: event.subject,
            start: event.start.dateTime,
            end: event.end.dateTime,
            bodyPreview: event.bodyPreview,
            joinUrl: event.joinUrl,
          }))}
          eventContent={eventContent}
        />
      </div>
    </div>
  );
};

export default EventsCalendar;
