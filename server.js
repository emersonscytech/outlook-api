const path = require('path');
const express = require('express');
const app = express();

const bodyParser = require('body-parser');
const moment = require('moment');

const cookieParser = require('cookie-parser');
const querystring = require('querystring');
const outlook = require('node-outlook');
const authHelper = require('./outlook/authHelper');

app.use(bodyParser.json());

app.use(cookieParser());

const tokenReceived = (req, res, error, token) => {
	if (error) {
		console.log('ERROR getting token:'  + error);
		res.send('ERROR getting token: ' + error);
	} else {
		app.set("access_token", token.token.access_token);
		app.set("refresh_token", token.token.refresh_token);
		app.set("email", authHelper.getEmailFromIdToken(token.token.id_token));
		res.redirect('/logincomplete');
	}
}

app.set('views', path.join(__dirname, '/client/html'));
app.engine('html', require('ejs').renderFile);
app.set('view engine', 'html');
app.use(express.static(path.join(__dirname, 'client')));
app.use('/node_modules', express.static( path.join(__dirname, './node_modules') ));

app.get("/", (req, res, next) => {
	res.render("index");
});

app.get('/auth', (req, res) => {
	res.send(authHelper.getAuthUrl());
});

app.get('/auth/allow', (req, res) => {
	const authCode = req.query.code;

	if (authCode) {
		console.log('');
		console.log('Retrieved auth code in /authorize: ' + authCode);
		authHelper.getTokenFromCode(authCode, tokenReceived, req, res);
	} else {
		console.log('/authorize called without a code parameter, redirecting to login');
		res.redirect('/');
	}
});

app.get('/logincomplete', (req, res) => {
	const accessToken = app.get("access_token");
	const refreshToken = app.get("access_token");
	const email = app.get("email");

	if (accessToken === undefined || refreshToken === undefined) {
		console.log('/logincomplete called while not logged in');
		res.redirect('/');
		return;
	}

	res.redirect("http://localhost:3000/#!/calendar");
});


app.get('/refreshtokens', (req, res) => {
	const refreshToken = app.get("access_token");

	if (refreshToken === undefined) {
		console.log('no refresh token in app');
		res.redirect('/');
	} else {
		authHelper.getTokenFromRefreshToken(refreshToken, tokenReceived, req, res);
	}
});

app.get('/logout', (req, res) => {
	app.set("access_token", null);
	app.set("refresh_token", null);
	app.set("email", null);
	res.redirect('/');
});

app.get('/calendar', (req, res) => {
	const accessToken = app.get("access_token");
	const email = app.get("email");

	if (accessToken === undefined || email === undefined) {
		console.log('/calendar called while not logged in');
		res.json({
			error: "/calendar called while not logged in",
			response: []
		});
	}

	outlook.base.setApiEndpoint('https://outlook.office.com/api/v2.0');
	outlook.base.setAnchorMailbox(app.get("email"));
	outlook.base.setPreferredTimeZone('Eastern Standard Time');

	let requestUrl = app.get("syncUrl");

	if (requestUrl === undefined) {
		requestUrl = outlook.base.apiEndpoint() + '/Me/CalendarView';
	}

	const startDate = moment().startOf('day');
	const endDate = moment(startDate).add(7, 'days');

	const params = {
		startDateTime: startDate.toISOString(),
		endDateTime: endDate.toISOString()
	};

	const headers = {
		Prefer: [
			//'odata.track-changes',
			'odata.maxpagesize=5'
		]
	};

	const apiOptions = {
		url: requestUrl,
		token: accessToken,
		headers: headers,
		query: params
	};

	// outlook.calendar.getEvents() // future
	outlook.base.makeApiCall(apiOptions, (error, response) => {
		if (error) {
			console.log(JSON.stringify(error));
			res.send(JSON.stringify(error));
		} else {
			if (response.statusCode !== 200) {
				res.json({
					"error": 'API Call returned ' + response.statusCode
				});
			} else {
				const nextLink = response.body['@odata.nextLink'];
				const deltaLink = response.body['@odata.deltaLink'];

				if (nextLink !== undefined) {
					app.set("syncUrl", nextLink);
				}

				if (deltaLink !== undefined) {
					app.set("syncUrl", deltaLink);
				}

				res.json({
					"email": email,
					"response": response.body.value
				});
			}
		}
	});
});

app.post('/calendar/create', (req, res) => {
	const eventId = req.params.eventId;
	const accessToken = app.get("access_token");

	if (accessToken === undefined) {
		res.json({
			error: "Token not found."
		});
	}

	const eventParameters = {
		token: accessToken,
		event: req.body
	};

	outlook.calendar.createEvent(eventParameters, (error, event) => {
		if (error) {
			res.json({
				"error": error
			});
		} else {
			res.json({
				"event": event
			});
		}
	});
});

app.get('/calendar/:eventId', (req, res) => {
	const eventId = req.params.eventId;
	const accessToken = app.get("access_token");
	const email = app.get("email");

	if (eventId === undefined || accessToken === undefined) {
		res.json({
			error: `${eventId} not found.`
		});
		return;
	}

	const select = {
		'$select': 'Subject,Attendees,Location,Start,End,IsReminderOn,ReminderMinutesBeforeStart'
	};

	const getEventParameters = {
		token: accessToken,
		eventId: eventId,
		odataParams: select
	};

	outlook.calendar.getEvent(getEventParameters, function(error, event) {
		if (error) {
			console.log("API ERROR: ", error);
			res.json({
				error: error
			});
		} else {
			res.json({
				"email": email,
				"event": event
			});
		}
	});
});

app.put('/calendar/update/:eventId', (req, res) => {
	const eventId = req.params.eventId;
	const accessToken = app.get("access_token");

	if (eventId === undefined || accessToken === undefined) {
		res.json({
			error: `${eventId} not found.`
		});
	}

	const reqData = req.body;

	[
		"@odata.context",
		"@odata.id",
		"@odata.etag",
		"Calendar@odata.associationLink",
		"Calendar@odata.navigationLink"
	].forEach((removingOption) => {
		if (reqData[removingOption]) {
			delete reqData[removingOption];
		}
	});

	const updateEventParameters = {
		token: accessToken,
		eventId: eventId,
		update: reqData
	};

	outlook.calendar.updateEvent(updateEventParameters, (error, event) => {
		if (error) {
			console.log("API ERROR: ", error);
			res.json({
				error: error
			});
		} else {
			res.json({
				"event": event
			});
		}
	});
});

app.delete('/calendar/:eventId', (req, res) => {
	const eventId = req.params.eventId;
	const accessToken = app.get("access_token");

	if (eventId === undefined || accessToken === undefined) {
		res.json({
			error: `${eventId} not found.`
		});
		return;
	}

	const deleteEventParameters = {
		token: accessToken,
		eventId: eventId
	};

	outlook.calendar.deleteEvent(deleteEventParameters, (error, event) => {
		if (error) {
			console.log("API ERROR: ", error);
			res.json({
				"status": error
			});
		} else {
			res.json({
				"status": "deleted"
			});
		}
	});
});

const server = app.listen(3000, () => {
	const host = server.address().address;
	const port = server.address().port;
	console.log('Example app listening at http://%s:%s', host, port);
});