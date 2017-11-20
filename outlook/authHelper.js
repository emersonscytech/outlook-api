let configuration;
try {
	configuration = require("../configuration.json");
} catch(e) {
	console.error("Please create a file named configuration.json in current folder!");
	console.error("Read configuration-template.json.\n");
	process.exit(-1);
}

const clientId = configuration.id;
const clientSecret = configuration.secret;
const redirectUri = configuration.redirectUri;

const scopes = [
	'openid',
	'profile',
	'offline_access',
	'https://outlook.office.com/calendars.readwrite'
];

const credentials = {
	client: configuration.client,
	auth: {
		tokenHost: 'https://login.microsoftonline.com',
		authorizePath: 'common/oauth2/v2.0/authorize',
		tokenPath: 'common/oauth2/v2.0/token'
	}
};

const oauth2 = require('simple-oauth2').create(credentials);

module.exports = {
	getAuthUrl: () => {
		const returnVal = oauth2.authorizationCode.authorizeURL({
			redirect_uri: redirectUri,
			scope: scopes.join(' ')
		});
		return returnVal;
	},

	getTokenFromCode: (auth_code, callback, request, response) => {
		oauth2.authorizationCode.getToken({
			code: auth_code,
			redirect_uri: redirectUri,
			scope: scopes.join(' ')
		}, (error, result) => {
			if (error) {
				console.log('Access token error: ', error.message);
				callback(request ,response, error, null);
			} else {
				const token = oauth2.accessToken.create(result);
				callback(request, response, null, token);
			}
		});
	},

	getEmailFromIdToken: (idToken) => {
		const tokenParts = idToken.split('.');
		const encodedToken = new Buffer(tokenParts[1].replace('-', '+').replace('_', '/'), 'base64');
		const decodedToken = encodedToken.toString();
		const jwt = JSON.parse(decodedToken);
		return jwt.preferred_username
	},

	getTokenFromRefreshToken: (refresh_token, callback, request, response) => {
		const token = oauth2.accessToken.create({
			refresh_token: refresh_token,
			expires_in: 0
		});

		token.refresh((error, result) => {
			if (error) {
				console.log('Refresh token error: ', error.message);
				callback(request, response, error, null);
			} else {
				console.log('New token: ', result.token);
				callback(request, response, null, result);
			}
		});
	}
};
