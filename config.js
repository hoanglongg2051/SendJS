/*
HOW TO USE
- npm install colors figlet nodemailer fs fs-extra yesno
- node main.js

RANDOM LIST
- ##email##
- ##username##
- ##image##
- ##link##
- ##blank_1## ( Blank char. 1 is the length )
- ##number_10##
- ##letter_10##
- ##letterup_10##
- ##letterlow_10##
- ##letternumber_10##
- ##letternumberup_10##
- ##letternumberlow_10##
- ##32bit_10## ( 32bit character, such as md5 / sha1 )
- Costumize your custom random in customrandom.js
*/


exports.send = {

	SMTP : {
		File 	  : "Smtp/smtp.txt", // Smtp file
		Host 	  : "smtp.office365.com", // Smtp host
		Port 	  : 587, // Smtp port
		Hostname  : "", // Set local hostname ( Levae it blank if not using )
	},

	LIST : {
		File 	  : "List/test.txt", // List file
		RemoveDup : false, // Remove list duplicate ( true/false )
		Logfailed : false // Log failed send ( Stored in Log/failed.txt )
	},

	LETTER : {
		Folder 	  : "Letter", // Letter folder
		Type 	  : "text", // Letter type ( html/text )
		Encoding  : "utf-8", // Letter encoding ( utf-8/hex/base64 )
	},

	LINK : {
		File 	  : "Link/link.txt", // Link file
	},

	ALTERNATIVE : {
		File 	  : "", // Alternative file ( it wil shown as multipart "Not work in text mode" ) ( Leave it blank if not using )
		Encoding  : "quoted-printable" // Alternative text encoding ( quoted-printable/base64 )
	},

	ATTACHMENT : {
		Folder 	  : "", // Attachment Folder ( Leave it blank if not using )
		Name 	  : "", // Attachment custom name
		Encoding  : "base64", // Attachment encoding ( base64/hex/binary )
	},

	IMAGE	 : {
		File 	  : "", // Image File ( Leave it blank if not using ) ( Usage : <img src="##image##"> )
		Name 	  : "Logo ##letternumberup_5##.png", // Image custom name
		Encoding  : "base64" // Image encoding ( base64/hex/binary )
	},

	EXPERIMENT : {
		Emailtest : "", // Email for test ( Levae it blank if not using )
		Testafter : 1000 // Test after x number of sent
	},

	SEND : {
		Delay 	  : 750, // Sending delay for x miliseconds
		Priority  : "", // Email priority ( high/normal/low ) ( Levae it blank if not using )
		Returnpath: "", // Email return path ( Levae it blank if not using )
		Messageid : "", // Custom message id ( Levae it blank if not using )

		Bccmode   : false, // Bcc mode ( true/false )
		Bccto	  : "", // Add to while using bcc mode

		Fromname  : "", // From name ( Levae it blank if not using )
		Frommail  : "", // From mail ( Levae it blank if not using )
		Subject   : "", // Subject ( Levae it blank if not using )
		Replto 	  : "" // Reply-to ( Levae it blank if not using )
	},

	HEADER : {
		Useheader : false,
		Headerval : {
			"x-job" : "##number_9##_##number_6##",
			"Feedback-ID" : "##number_9##:##number_6##:13.111.103.197:sfmktgcld"
		}
	}
};