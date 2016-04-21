/*
 * Dieses Skript hilft beim Archivieren von E-Mails.
 *
 * Verwendung:
 * Eine oder mehrere Outlook E-Mail-Dateien (.msg) aus dem Dateimanager auf
 * diese Datei ziehen (direkt aus Outlook heraus klappt nicht).
 *
 * Ergebnis:
 * Die E-Mails werden in dieses Verzeichnis verschoben oder umbenannt, dem
 * Dateinamen wird das Datum der E-Mail vorangestellt. Existiert eine Datei mit
 * identischem Dateinamen bereits, kommt eine fortlaufende Nummer zum Einsatz.
 * Ist dies jedoch die gleiche E-Mail, wird sie nicht noch einmal gespeichert
 * und ein Hinweis ausgegeben.
 */

// Activate workarounds for testing with Wine (1.9.8 used), keep at false for
// normal operation.
var winetest = false;

// Testing without installed Microsoft Outlook, keep at false for normal
// operation.
var nooutlooktest = false;

var fso = new ActiveXObject("Scripting.FileSystemObject");
var ol;

if (nooutlooktest) {
	function OutlookApplication()
	{
		this.CreateItemFromTemplate = function(TemplatePath) {
			if (!fso.FileExists(TemplatePath))
				return null;

			return {
				SentOn: new Date("Tue, 22 Mar 2016 09:57:55 GMT+0100"),
				SenderEmailAddress: "<name@domain>",
				//SenderEmailAddress: TemplatePath,
				Close: function(SaveMode) {}
			};
		}
	}
	ol = new OutlookApplication;
} else {
	ol = new ActiveXObject("Outlook.Application");
}

// Outlook.Application constants
var olDiscard = 1;

/*
 * Open the Outlook email given by "path". Returns an Outlook MailItem object,
 * null on error.
 */
function openOutlookEmailFile(path)
{
	try {
		return ol.CreateItemFromTemplate(path);
	} catch (exception) {
		return null;
	}
}

/*
 * Read value of key from the Outlook email. Returns different types.
 */
function getEmailPropVal(itemorpath, key)
{
	var mailitem;
	var val = null;

	// Handle variable parameter type
	if (typeof itemorpath === "string") {
		// Path
		mailitem = openOutlookEmailFile(itemorpath);

		if (!mailitem)
			return null;
	} else {
		// Outlook MailItem
		mailitem = itemorpath;
	}

	if (key === "Date") {
		// Outlook Date (typeof returns "date" instead of "object" as
		// for the normal JavaScript/JScript Date object)
		val = mailitem.SentOn;
	} else if (key === "From") {
		// String
		val = mailitem.SenderEmailAddress
	}

	if (typeof itemorpath === "string")
		mailitem.Close(olDiscard);

	return val;
}

/*
 * Read the email "Date" property value. Returns a Date object, null if
 * invalid.
 */
function getOutlookEmailDate(itemorpath)
{
	var oldate, date;

	oldate = getEmailPropVal(itemorpath, "Date");
	date = new Date(oldate);

	if (isNaN(date))
		return null;

	return date;
}

/*
 * Check whether two emails are the same, the required "From" and "Date" fields
 * have to be equal therefore (the optional "Message-ID" field cannot be read
 * via Outlook MailItem interface). Returns a Boolean, false on error.
 */
function isSameEmail(paths)
{
	var from = [], date = [];

	for (var i = 0; i < 2; i++) {
		from[i] = getEmailPropVal(paths[i], "From");
		date[i] = getOutlookEmailDate(paths[i]);

		// Make the date comparable
		if (!date[i])
			return false;
		date[i] = date[i].getTime();
	}

	if (from[0] && from[1])
		return from[0] === from[1] && date[0] === date[1];
	else
		return false;
}

/*
 * Add a leading 0 if the value is a single digit number. Returns a String.
 */
function pad(val)
{
	return val < 10 ? "0" + val : val.toString();
}

/*
 * Prepend a date prefix "m_YYYY-MM-DD_" to the email filename, time zone is
 * the local zone, and use "destdir" as directory. If the resulting file
 * already exists, append a sequential number " (n)" with n = 2 or higher to
 * the email base filename. Returns a String, this is a filename of an existing
 * file if it's the same email, null if limit of 999 files exceeded.
 */
function getDestPath(srcpath, destdir, date)
{
	var srcfilename = fso.GetFileName(srcpath);

	// Cut a possibly existing date prefix from the filename to avoid
	// prepending a second date prefix. This would happen if using this
	// script with already renamed emails.
	var destfilenamepart = srcfilename.replace(/^m_[0-9]{4}-[01][0-9]-[0-3][0-9]_/, "");

	var destpath = destdir
		+ "\\m_"
		+ date.getFullYear()
		+ "-"
		+ pad(date.getMonth() + 1)
		+ "-"
		+ pad(date.getDate())
		+ "_"
		+ destfilenamepart;

	// Search for an alternative filename if the file already exists (as a
	// different email).

	var base = fso.GetParentFolderName(destpath) + "\\" + fso.GetBaseName(destpath);
	var ext = fso.GetExtensionName(destpath) ? "." + fso.GetExtensionName(destpath) : "";
	var num = 2;

	while (fso.FileExists(destpath)) {
		if (isSameEmail([srcpath, destpath]))
			break;

		// Artificial limit
		if (num > 999)
			return null;

		destpath = base + " (" + num + ")" + ext;
		num++;
	}

	return destpath;
}

/*
 * Save (move) an Outlook email (.msg file) from file "srcpath" with adapted
 * filename to "destdir". Returns a user message String on failure.
 */
function saveEmail(srcpath, destdir)
{
	var mailitem;
	var emaildate;
	var destpath;

	// Read email date from source file

	mailitem = openOutlookEmailFile(srcpath);

	if (!mailitem)
		return "Datei konnte nicht geoeffnet werden.";

	emaildate = getOutlookEmailDate(mailitem);

	mailitem.Close(olDiscard);

	if (!emaildate)
		return "Kein gueltiges Datum gefunden. Ist dies eine Outlook E-Mail (.msg-Datei)?";

	// Determine destination filename

	destpath = getDestPath(srcpath, destdir, emaildate);

	if (!destpath)
		return "Obergrenze von 999 gleichen Dateinamen erreicht.";

	if (fso.FileExists(destpath))
		return "E-Mail existiert bereits: \"" + fso.GetFileName(destpath) + "\"";

	// Move source to destination file

	try {
		if (winetest) {
			// FileSystemObject.MoveFile() not implemented
			fso.CopyFile(srcpath, destpath, false);
			fso.DeleteFile(srcpath);
		} else {
			fso.MoveFile(srcpath, destpath);
		}
	} catch (exception) {
		// JScript runtime error, e.g.:
		// 800A0035: Die Datei wurde nicht gefunden.
		// 800A003A: Die Datei ist bereits vorhanden.
		// 800A0046: Erlaubnis verweigert.
		return "Fehlende Berechtigung oder anderes Problem.";
	}
}

/*
 * Iterate through all parameters, each one representing an email file path,
 * and save (move) it with adapted filename to the directory of the script.
 * Collect all occuring problems for a single error feedback at the end.
 */
function main()
{
	var argv = WScript.Arguments;
	var destdir = fso.GetParentFolderName(WScript.ScriptFullName);
	var error, errorcnt = 0, feedback = "";

	for (var i = 0; i < argv.length; i++) {
		error = saveEmail(argv(i), destdir);

		if (error) {
			errorcnt++;
			feedback += "\"" + argv(i) + "\":\n" + error + "\n\n";
		}
	}

	if (feedback)
		WScript.echo(errorcnt + " Dateien konnten nicht verarbeitet werden:\n\n" + feedback);
}

main();
