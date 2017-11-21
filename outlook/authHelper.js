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

const getAuthUrl = () => {
	const returnVal = oauth2.authorizationCode.authorizeURL({
		redirect_uri: redirectUri,
		scope: scopes.join(' ')
	});
	return returnVal;
};

const getTokenFromCode = (authCode, errorCallback, successCallback) => {
	const options = {
		code: authCode,
		redirect_uri: redirectUri,
		scope: scopes.join(' ')
	}
	oauth2.authorizationCode.getToken(options, (error, result) => {
		if (error) {
			errorCallback(error);
		} else {
			const token = oauth2.accessToken.create(result);
			successCallback(token);
		}
	});
}

const getEmailFromIdToken = (idToken) => {
	const tokenParts = idToken.split('.');
	const encodedToken = new Buffer(tokenParts[1].replace('-', '+').replace('_', '/'), 'base64');
	const decodedToken = encodedToken.toString();
	const jsonToken = JSON.parse(decodedToken);

	return jsonToken.preferred_username
}

const getTokenFromRefreshToken = (refreshToken, errorCallback, successCallback) => {

	const token = oauth2.accessToken.create({
		refresh_token: refreshToken,
		expires_in: 0
	});

	token.refresh((error, result) => {
		if (error) {
			errorCallback(error);
		} else {
			successCallback(result);
		}
	});
};

module.exports = {
	getAuthUrl,
	getTokenFromCode,
	getEmailFromIdToken,
	getTokenFromRefreshToken,
};
