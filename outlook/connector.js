const url = require('url');
const authHelper = require('./authHelper');
const microsoftGraph = require("@microsoft/microsoft-graph-client");

const mail = (response, request) => {
	getAccessToken(request, response, function(error, token) {
		console.log('Token found in cookie: ', token);

		const email = getValueFromCookie('node-tutorial-email', request.headers.cookie);
		console.log('\nEmail found in cookie: ', email);

		if (token) {
			response.writeHead(200, {'Content-Type': 'text/html; charset=UTF-8'});
			response.write('<div><h1>Your inbox</h1></div>');

			const client = microsoftGraph.Client.init({ // Create a Graph client
				authProvider: (done) => {
					done(null, token); // Just return the token
				}
			});

			// Get the 10 newest messages
			client
				.api('/me/mailfolders/inbox/messages')
				.header('X-AnchorMailbox', email)
				.top(10)
				.select('subject,from,receivedDateTime,isRead')
				.orderby('receivedDateTime DESC')
				.get((err, apiResponse) => {
					if (err) {
						console.log('getMessages returned an error: ' + err);
						response.write('<p>ERROR: ' + err + '</p>');
						response.end();
					} else {
						console.log('getMessages returned ' + apiResponse.value.length + ' messages.');
						response.write('<table><tr><th>From</th><th>Subject</th><th>Received</th></tr>');

						apiResponse.value.forEach(message => {
							console.log('  Subject: ' + message.subject);
							const from = message.from ? message.from.emailAddress.name : 'NONE';
							response.write(
								'<tr><td>' + from + '</td><td>'
								+ (message.isRead ? '' : '<b>') + message.subject + (message.isRead ? '' : '</b>') +
								'</td><td>' + message.receivedDateTime.toString() + '</td></tr>'
							);
						});

						response.write('</table>');
						response.end();
					}
				});
		} else {
			response.writeHead(200, {
				'Content-Type': 'text/html; charset=UTF-8'
			});
			response.write('<p> No token found in cookie!</p>');
			response.end();
		}
	});
}

const calendar = (response, request) => {
	getAccessToken(request, response, (error, token) => {
		const email = getValueFromCookie('node-tutorial-email', request.headers.cookie);

		if (token) {
			response.writeHead(200, {
				'Content-Type': 'text/html; charset=UTF-8'
			});
			response.write('<div><h1>Your calendar</h1></div>');

			const client = microsoftGraph.Client.init({
				authProvider: (done) => {
					done(null, token);
				}
			});

			client
				.api('/me/events')
				.header('X-AnchorMailbox', email)
				.top(10)
				.select('subject,start,end')
				.orderby('start/dateTime DESC')
				.get((err, res) => {
					if (err) {
						console.log('getEvents returned an error: ' + err);
						response.write('<p>ERROR: ' + err + '</p>');
						response.end();
					} else {
						console.log('getEvents returned ' + res.value.length + ' events.');
						response.write('<table><tr><th>Subject</th><th>Start</th><th>End</th><th>Attendees</th></tr>');
						res.value.forEach((event) => {
							console.log('  Subject: ' + event.subject);
							response.write('<tr><td>' + event.subject + '</td><td>' + event.start.dateTime.toString() + '</td><td>' + event.end.dateTime.toString() + '</td></tr>');
						});

						response.write('</table>');
						response.end();
					}
				});
		} else {
			response.writeHead(200, {'Content-Type': 'text/html; charset=UTF-8'});
			response.write('<p> No token found in cookie!</p>');
			response.end();
		}
	});
}

const contacts = (response, request) => {
	getAccessToken(request, response, (error, token) => {
		const email = getValueFromCookie('node-tutorial-email', request.headers.cookie);

		if (token) {
			response.writeHead(200, {'Content-Type': 'text/html; charset=UTF-8'});
			response.write('<div><h1>Your contacts</h1></div>');

			const client = microsoftGraph.Client.init({
				authProvider: (done) => {
					done(null, token);
				}
			});

			client
				.api('/me/contacts')
				.header('X-AnchorMailbox', email)
				.top(10)
				.select('givenName,surname,emailAddresses')
				.orderby('givenName ASC')
				.get((err, res) => {
					if (err) {
						console.log('getContacts returned an error: ' + err);
						response.write('<p>ERROR: ' + err + '</p>');
						response.end();
					} else {
						console.log('getContacts returned ' + res.value.length + ' contacts.');
						response.write('<table><tr><th>First name</th><th>Last name</th><th>Email</th></tr>');

						res.value.forEach((contact) => {
							const email = contact.emailAddresses[0] ? contact.emailAddresses[0].address : 'NONE';
							response.write('<tr><td>' + contact.givenName + '</td><td>' + contact.surname + '</td><td>' + email + '</td></tr>');
						});

						response.write('</table>');
						response.end();
					}
				});
		} else {
			response.writeHead(200, {'Content-Type': 'text/html; charset=UTF-8'});
			response.write('<p> No token found in cookie!</p>');
			response.end();
		}
	});
}

const home = (response, request) => {
	console.log('Request handler \'home\' was called.');
	response.writeHead(200, {
		'Content-Type': 'text/html'
	});
	response.write('<p>Please <a href="' + authHelper.getAuthUrl() + '">sign in</a> with your Office 365 or Outlook.com account.</p>');
	response.end();
}

const getUserEmail = (token, callback) => {
	var client = microsoftGraph.Client.init({
		authProvider: (done) => {
			done(null, token);
		}
	});

	client
		.api('/me')
		.get((err, res) => {
			if (err) {
				callback(err, null);
			} else {
				callback(null, res.userPrincipalName);
			}
		});
}

const authorize = (response, request) => {
	console.log('Request handler \'authorize\' was called.');

	var url_parts = url.parse(request.url, true); // The authorization code is passed as a query parameter
	var code = url_parts.query.code;
	console.log('Code: ' + code);
	authHelper.getTokenFromCode(code, tokenReceived, response);
}

const tokenReceived = (response, error, token) => {
	if (error) {
		console.log('Access token error: ', error.message);
		response.writeHead(200, {'Content-Type': 'text/html'});
		response.write('<p>ERROR: ' + error + '</p>');
		response.end();
	} else {
		getUserEmail(token.token.access_token, (error, email) => {
			if (error) {
				console.log('getUserEmail returned an error: ' + error);
				response.write('<p>ERROR: ' + error + '</p>');
				response.end();
			} else if (email) {
				const cookies = ['node-tutorial-token=' + token.token.access_token + ';Max-Age=4000',
				'node-tutorial-refresh-token=' + token.token.refresh_token + ';Max-Age=4000',
				'node-tutorial-token-expires=' + token.token.expires_at.getTime() + ';Max-Age=4000',
				'node-tutorial-email=' + email + ';Max-Age=4000'];

				response.setHeader('Set-Cookie', cookies);

				response.writeHead(301, {
					'Location': 'http://localhost:8000/mail'
				});

				response.end();
			} else {
				console.log("ELSE!", email)
			}
		});
	}
}

const getAccessToken = (request, response, callback) => {
	var expiration = new Date(parseFloat(getValueFromCookie('node-tutorial-token-expires', request.headers.cookie)));

	if (expiration <= new Date()) {
		console.log('TOKEN EXPIRED, REFRESHING'); // refresh token
		const refresh_token = getValueFromCookie('node-tutorial-refresh-token', request.headers.cookie);
		authHelper.refreshAccessToken(refresh_token, function(error, newToken){
			if (error) {
				callback(error, null);
			} else if (newToken) {
				const cookies = ['node-tutorial-token=' + newToken.token.access_token + ';Max-Age=4000',
				'node-tutorial-refresh-token=' + newToken.token.refresh_token + ';Max-Age=4000',
				'node-tutorial-token-expires=' + newToken.token.expires_at.getTime() + ';Max-Age=4000'];
				response.setHeader('Set-Cookie', cookies);
				callback(null, newToken.token.access_token);
			}
		});
	} else {
		const access_token = getValueFromCookie('node-tutorial-token', request.headers.cookie);
		callback(null, access_token); // Return cached token
	}
}

const getValueFromCookie = (valueName, cookie) => {
	if (cookie.indexOf(valueName) !== -1) {
		const start = cookie.indexOf(valueName) + valueName.length + 1;
		let end = cookie.indexOf(';', start);
		end = end === -1 ? cookie.length : end;
		return cookie.substring(start, end);
	}
}

module.exports.getAccessToken = getAccessToken;