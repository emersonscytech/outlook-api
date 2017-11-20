interface CalendarResponse {
	response: EventResponse[],
	error: string,
	email: string
}

interface EventDate {
	DateTime?: Date | string,
	TimeZone?: string
}

interface EventLocation {
	DisplayName: string,
	Address: Object,
	Coordinates: Object
}

interface EventResponse {
	event: OutlookEvent,
	email: string,
	error: string
}

interface OutlookEvent {
	Id?: string,
	IsReminderOn?: boolean,
	Subject: string,
	Body?: {
		ContentType?: string,
		Content?: string
	}
	Start?: EventDate,
	End?: EventDate,
	Location?: EventLocation,
}

interface OutlookEventDeleteResponse {
	status: string
}