class OutlookEventService {
	private $http: angular.IHttpService;

	public constructor($http: angular.IHttpService) {
		this.$http = $http;
	}

	public createEvent(outlookEvent: OutlookEvent): angular.IHttpPromise<EventResponse> {
		return this.$http.post("calendar/create", outlookEvent);
	}

	public getEvent(eventId: string): angular.IHttpPromise<EventResponse> {
		return this.$http.get(`calendar/${eventId}`);
	}

	public updateEvent(outlookEvent: OutlookEvent): angular.IHttpPromise<EventResponse> {
		return this.$http.put(`calendar/update/${outlookEvent.Id}`, outlookEvent);
	}

	public deleteEvent(eventId: string): angular.IHttpPromise<OutlookEventDeleteResponse> {
		return this.$http.delete(`calendar/${eventId}`);
	}
}

class OutlookEventController {
	public outlookEvent: OutlookEvent;
	public eventId: string;
	public message: string;
	public isEditing = false;

	public constructor(protected service: OutlookEventService, $routeParams: angular.route.IRouteParamsService) {
		if ($routeParams['eventId']) {
			this.eventId = $routeParams['eventId'];
			this.isEditing = true;
			this.service.getEvent(this.eventId).then((response: angular.IHttpPromiseCallbackArg<EventResponse>) => {
				this.outlookEvent = response.data.event;
			});
		} else {
			this.outlookEvent = {
				"Subject": "",
				"Body": {
					"ContentType": "HTML",
				},
				"Start": {
					"TimeZone": "Eastern Standard Time"
				},
				"End": {
					"TimeZone": "Eastern Standard Time"
				}
			};
		}
	}

	public submit(): void {
		if (this.isEditing) {
			this.service.updateEvent(this.outlookEvent).then((response: angular.IHttpPromiseCallbackArg<EventResponse>) => {
				console.log(response.data.event);
				this.message = "The Event Was Updated!";
				setTimeout(() => {
					window.location.href = "#!/calendar/";
				}, 1500);
			});
		} else {
			this.service.createEvent(this.outlookEvent).then((response: angular.IHttpPromiseCallbackArg<EventResponse>) => {
				this.message = "The Event Was Created!";
				console.log(response.data.event);
				setTimeout(() => {
					window.location.href = "#!/calendar/";
				}, 1500);
			});
		}
	}

	public deleteEvent(): void {
		if (confirm(`Do you really want to remove the event "${this.outlookEvent.Subject}" ?`)) {
			this.service.deleteEvent(this.outlookEvent.Id).then((response: angular.IHttpPromiseCallbackArg<OutlookEventDeleteResponse>) => {
				this.message = response.data.status;
				setTimeout(() => {
					window.location.href = "#!/calendar/";
				}, 1500);
			});
		}
	}
}

angular.module("outlook").service("OutlookEventService", ["$http", OutlookEventService]);
angular.module('outlook').controller('OutlookEventController', ['OutlookEventService', '$routeParams', OutlookEventController]);
