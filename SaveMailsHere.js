/*
 * Dieses Skript hilft beim Archivieren von E-Mails.
 *
 * Verwendung:
 * Eine oder mehrere E-Mails aus dem E-Mail-Programm oder Dateimanager auf
 * diese Datei ziehen.
 *
 * Ergebnis:
 * Die E-Mails werden in dieses Verzeichnis gespeichert, dem Dateinamen wird
 * das Datum der E-Mail vorangestellt. Existiert eine Datei mit identischem
 * Dateinamen bereits, kommt eine fortlaufende Nummer zum Einsatz. Ist dies
 * jedoch die gleiche E-Mail, wird sie nicht noch einmal gespeichert und ein
 * Hinweis ausgegeben.
 */

// Activate workarounds for testing with Wine (1.9.8 used), keep at false for
// normal operation.
var winetest = false;

var fso = new ActiveXObject("Scripting.FileSystemObject");

// Scripting.FileSystemObject constants
var ForReading = 1;
var TristateFalse = 0;

/*
 * Open the email given by "path". Returns a TextStream object, null on error.
 */
function openEmailFile(path)
{
	try {
		// Open read-only as ASCII
		return fso.OpenTextFile(path, ForReading, false, TristateFalse);
	} catch (exception) {
		// JScript runtime error, e.g.:
		// 800A0035: Die Datei wurde nicht gefunden.
		return null;
	}
}

/*
 * Read value of key from the email header field. Returns a String, null if not
 * found.
 */
function getEmailHeaderVal(streamorpath, key)
{
	var stream, line, regex;
	var keymatch, val = null;

	var regex = new RegExp("^" + key + ": (.*)");

	// Handle variable parameter type
	if (typeof streamorpath === "string") {
		// Path
		stream = openEmailFile(streamorpath);

		if (!stream)
			return null;
	} else {
		// Stream
		stream = streamorpath;
	}

	// Search the whole email header (until blank line)
	while (!stream.AtEndOfStream && line !== "") {
		if (winetest) {
			// TextStream.Readline() not implemented
			if (key === "Date") {
				//line = "";
				//line = "Date: ";
				//line = "Date: garbage";
				//line = "Date: Tue, 35 Mar 2016 09:57:55 +0100";
				line = "Date: Tue, 22 Mar 2016 09:57:55 +0100";
			} else if (key === "Message-ID") {
				//line = "";
				//line = "Message-ID: ";
				line = "Message-ID: <unique-id@domain>";
			} else if (key === "From") {
				//line = "";
				//line = "From: ";
				line = "From: \"Name\" <name@domain>";
			} else {
				line = key + ": dummy value";
			}
		} else {
			line = stream.ReadLine();
		}

		keymatch = line.match(regex);

		if (keymatch) {
			val = keymatch[1];
			break;
		}
	}

	if (typeof streamorpath === "string")
		stream.Close();

	return val;
}

/*
 * Read the email "Date" field value. Returns a Date object, null if not found
 * or invalid.
 */
function getEmailDate(streamorpath)
{
	var datestr, date;

	datestr = getEmailHeaderVal(streamorpath, "Date");

	// Abort if "Date" field is not found (would result in 1970-01-01) or
	// an empty string (would result in NaN).
	if (!datestr)
		return null;

	if (winetest) {
		// Workaround for bug in Date.parse()
		datestr = datestr.replace(" +", " GMT+").replace(" -", " GMT-");
	}

	date = new Date(datestr);

	if (isNaN(date))
		return null;

	return date;
}

/*
 * Check whether two emails are the same, either the optional "Message-ID" or
 * the required "From" and "Date" fields have to be equal therefore. Returns a
 * Boolean, false on error.
 */
function isSameEmail(paths)
{
	var id = [], from = [], date = [];

	for (var i = 0; i < 2; i++) {
		id[i]   = getEmailHeaderVal(paths[i], "Message-ID");
		from[i] = getEmailHeaderVal(paths[i], "From");
		date[i] = getEmailHeaderVal(paths[i], "Date");
	}

	if (id[0] && id[1])
		return id[0] === id[1];
	else if (from[0] && from[1] && date[0] && date[1])
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
 * Save (move) an email from file "srcpath" with adapted filename to "destdir".
 * Returns a user message String on failure.
 */
function saveEmail(srcpath, destdir)
{
	var filestream;
	var emaildate;
	var destpath;

	// Read email date from source file

	filestream = openEmailFile(srcpath);

	if (!filestream)
		return "Datei konnte nicht geoeffnet werden.";

	emaildate = getEmailDate(filestream);

	filestream.Close();

	if (!emaildate)
		return "Kein gueltiges Datum gefunden. Ist dies eine E-Mail?";

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
		//
		// Drag & drop from email client:
		// If MoveFile() fails, the temporary file "srcpath", which has
		// been created on-the-fly by the dragging application, will
		// not be removed necessarily. This left over file may have
		// influence on filenames in subsequent drag & drop operations.
		// But since drag & drop is also possible from file browser, it
		// must not be removed here because "srcpath" is not a
		// temporary, but the original file.
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
		WScript.echo(errorcnt + " von " + argv.length + " Dateien konnten nicht verarbeitet werden:\n\n" + feedback);
}

main();
