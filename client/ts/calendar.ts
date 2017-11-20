class CalendarService {
	private $http: angular.IHttpService;

	public constructor($http: angular.IHttpService) {
		this.$http = $http;
	}

	public getCalendarData(): angular.IHttpPromise<CalendarResponse> {
		return this.$http.get('calendar');
	}
}

class CalendarController {
	public calendarData: CalendarResponse;
	public constructor(protected service: CalendarService) {
		this.service.getCalendarData().then((response: angular.IHttpPromiseCallbackArg<CalendarResponse>) => {
			this.calendarData = response.data;
		});
	}
}

angular.module("outlook").service("CalendarService", ["$http", CalendarService]);
angular.module('outlook').controller('CalendarController', ['CalendarService', CalendarController]);
