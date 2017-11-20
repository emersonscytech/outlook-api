class AuthService {
	private $http: angular.IHttpService;

	public constructor($http: angular.IHttpService) {
		this.$http = $http;
	}

	public getAuthUri(): angular.IHttpPromise<string> {
		return this.$http.get('auth');
	}
}

class AuthController {
	public authUri: string;
	public constructor(protected auth: AuthService) {
		this.auth.getAuthUri().then((response: angular.IHttpPromiseCallbackArg<string>) => {
			this.authUri = response.data;
		})
	}
}

angular.module("outlook").service("AuthService", ["$http", AuthService]);
angular.module('outlook').controller('AuthController', ['AuthService', AuthController]);
