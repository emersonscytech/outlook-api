angular.module("outlook", ["ngRoute", "ngMaterial"]);

angular.module("outlook").config(["$routeProvider",
	($routeProvider: angular.route.IRouteProvider) => {
		$routeProvider.when("/auth", {
			templateUrl: "html/auth.html",
			controller: "AuthController",
			controllerAs: "authCtrl"
		});

		$routeProvider.when("/calendar", {
			templateUrl: "html/calendar.html",
			controller: "CalendarController",
			controllerAs: "calendarCtrl"
		});

		$routeProvider.when("/calendar/create", {
			templateUrl: "html/event.html",
			controller: "OutlookEventController",
			controllerAs: "eventCtrl"
		});

		$routeProvider.when("/calendar/:eventId", {
			templateUrl: "html/event.html",
			controller: "OutlookEventController",
			controllerAs: "eventCtrl"
		});

		$routeProvider.when("/contacts", {
			templateUrl: "html/contacts.html",
			controller: "HomeController",
			controllerAs: "homeCtrl"
		});

		$routeProvider.when("/mail", {
			templateUrl: "html/mail.html",
			controller: "HomeController",
			controllerAs: "homeCtrl"
		});
		$routeProvider.otherwise("/auth");
	}
]);

angular.module("outlook").config(["$mdThemingProvider",
	($mdThemingProvider: angular.material.IThemingProvider) => {
		const theme = $mdThemingProvider.theme("default");
		theme.primaryPalette("green");
		theme.accentPalette("light-green");
	}
]);
