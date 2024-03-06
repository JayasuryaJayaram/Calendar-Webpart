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
import { Popover } from "antd";

interface IFormattedEvent {
  subject: any;
  startDate: any;
  endDate: any;
  startTime: any;
  endTime: any;
  eventDate: any;
  bodyPreview?: any;
  joinUrl?: any;
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
      --fc-event-bg-color: transparent;
      --fc-event-border-color: transparent;
  }
  .popover {
    max-width: none !important;
    /* Ensure the popover does not have a max-width */
  }
  .popover-arrow {
    border-right-color: #fff !important;
  }

  :where(.css-1qhpsh8).ant-popover .ant-popover-inner {
    padding: 0px;
  }

  :where(.css-dev-only-do-not-override-1qhpsh8).ant-popover .ant-popover-inner {
    padding: 0px;
  }

  :where(.css-dev-only-do-not-override-1qhpsh8).ant-btn-default:not(:disabled):not(.ant-btn-disabled):hover {
    color: #000;
  }

  .fc-direction-ltr .fc-daygrid-event.fc-event-end {
    height: 27px;
  }

  .fc-h-event .fc-event-main {
    height: 24px;
  }

  @media screen and (max-width: 426px) {
    .fc .fc-toolbar {
      display: flex;
      flex-wrap: wrap;
      line-height: 45px;
    }
    .n_c_8474018e {
      padding: 0px;
    }
    .calendarAppMain_89723c8a {
      padding: 0px;
    }
  }
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
                bodyPreview: event.bodyPreview || "",
                daysOfWeek: event.recurrence?.pattern?.daysOfWeek?.map(
                  (day: string) => {
                    switch (day) {
                      case "sunday":
                        return 0;
                      case "monday":
                        return 1;
                      case "tuesday":
                        return 2;
                      case "wednesday":
                        return 3;
                      case "thursday":
                        return 4;
                      case "friday":
                        return 5;
                      case "saturday":
                        return 6;
                      default:
                        return -1;
                    }
                  }
                ),
                startRecur: event.recurrence?.range?.startDate,
                endRecur: event.recurrence?.range?.endDate, // Changed to endDate
              }))
            );
          });
      });
  }, [props.context.msGraphClientFactory]);
  console.log("events", events);

  const eventContent = (eventInfo: any) => {
    console.log("eventInfo", eventInfo);

    const event: any = events.find(
      (evt: any) => evt.subject === eventInfo.event.title
    ); // Find the event in the events state
    if (!event) return null;

    const formattedEvent: IFormattedEvent = {
      subject: event.subject,
      startDate: event.start.dateTime,
      startTime: new Date(
        new Date(event.start.dateTime + "Z").toISOString()
      ).toLocaleString("en-US", {
        hour: "numeric",
        minute: "numeric",
        hour12: true,
      }),
      endDate: event.end.dateTime,
      endTime: new Date(
        new Date(event.end.dateTime + "Z").toISOString()
      ).toLocaleString("en-US", {
        hour: "numeric",
        minute: "numeric",
        hour12: true,
      }),
      eventDate: new Date(event.start.dateTime).toString(),
      bodyPreview: event.bodyPreview,
      joinUrl: event.joinUrl,
    };

    console.log("formattedEvent", formattedEvent);

    const content = (
      <div className={styles.popoverBox}>
        <div className={styles.popheader}>
          <b>Calendar</b> - <span>{props.context.pageContext.user.email}</span>
        </div>
        <div className={styles.popBody}>
          <div className={styles.popheading}>
            <img
              src={require("../assets/Icon1.svg")}
              alt="Icon"
              className={styles.popoverIcon}
              style={{ visibility: "hidden" }}
            />
            <span
              className={styles.contentStyle}
              style={{ textAlign: "inherit" }}
            >
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
              )}, ${formattedEvent.eventDate.substring(4, 10)} ${
                formattedEvent.startTime
              } - ${formattedEvent.endTime}`}
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
              style={{ top: "1px" }}
            />
            <p className={styles.contentStyle}>{formattedEvent.bodyPreview}</p>
          </div>
          <div style={{ display: formattedEvent.joinUrl ? "flex" : "none" }}>
            <button className={styles.joinBtn}>
              <a href={formattedEvent.joinUrl} target="_blank">
                Join
              </a>
            </button>
          </div>
        </div>
      </div>
    );

    return (
      <Popover content={content} trigger={["click", "hover"]} placement="right">
        <button className={styles.popoverButton}>
          <span>{formattedEvent.startTime} </span>
          <b> {formattedEvent.subject}</b>
        </button>
      </Popover>
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
            daysOfWeek: event.daysOfWeek,
            startRecur: event.startRecur,
            endRecur: event.endRecur,
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
