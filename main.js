"use strict";

const colors 		= require("colors");
const figlet 		= require("figlet");
const nodemailer 	= require("nodemailer");
const fs 			= require('fs');
const fsExtra 		= require('fs-extra');
const yesno 		= require('yesno');

const config 		= require("./config.js");
const custom 	 	= require("./customrandom.js");
const smtplist 	 	= fs.readFileSync(config.send.SMTP.File).toString().split("\n");
const list 	 		= fs.readFileSync(config.send.LIST.File).toString().split("\n");
const letters 		= fs.readdirSync(config.send.LETTER.Folder);
const links 	 	= fs.readFileSync(config.send.LINK.File).toString().split("\n");

console.log("\n");
console.log(
	colors.rainbow(
		figlet.textSync(" Office365", {
			font: "Doom",
			horizontalLayout: "default",
			verticalLayout: "default"
		})
	)
);
console.log(colors.rainbow(" =============================================="));
console.log(colors.brightRed(" +    Special Designed For Office365 SMTP     +"));
console.log(colors.brightRed(" +   Powered by : PT Pertamina Persero .Tbk   +"));
console.log(colors.rainbow(" =============================================="));
console.log("\n");

const attachments = (() => {
  if (config.send.ATTACHMENT.Folder) {
    return fs.readdirSync(config.send.ATTACHMENT.Folder);
  } else {
    return [];
  }
})();

function customrandom(input)
{
	var seed = custom.customrandom;
	Object.entries(seed).forEach(entry => {
		var [key, value] = entry;
		value = value.split("|");
		value = value[Math.floor(Math.random() * value.length)];
		input = input.replace(new RegExp(key, "g"), value);
	});

	return input;
}

function randomstring(type, length)
{
	var seed 	= "";
    var result 	= "";

	switch(type)
	{
	    case "number":
	        seed = "0123456789";
	        break;

	    case "letter":
	        seed = "abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ";
	        break;

	    case "letterup":
	        seed = "ABCDEFGHIJKLMNOPQRSTUVWXYZ";
	        break;

	    case "letterlow":
	        seed = "abcdefghijklmnopqrstuvwxyz";
	        break;

	    case "letternumber":
	        seed = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz01234567890123456789";
	        break;

	    case "letternumberup":
	        seed = "ABCDEFGHIJKLMNOPQRSTUVWXYZ01234567890123456789";
	        break;

	    case "letternumberlow":
	        seed = "abcdefghijklmnopqrstuvwxyz01234567890123456789";
	        break;

	    case "32bit":
	        seed = "abcdef01234567890123456789";
	        break;

	    case "blank":
	    	seed = ["‍", "͏", "⁪", "⁫", "‌", "⁬", "⁭", "⁮", "⁯"];
	    	break;
	}

    for ( var i = 0; i < length; i++ ) {
        if (type == "blank") {
			var random = Math.floor(Math.random() * seed.length);
			result += seed[random];
        } else {
        	result += seed.charAt(Math.floor(Math.random() * seed.length));
        }
    }
    return result;
}

function insertAfter(arr1, value, afterElement)
{
  	var result = [];
  	for(var i = 1; arr1.length > 0;i++) {
    	if(arr1.length > 0) {
      		result.push(arr1.shift());
    	}
    	if(i % afterElement == 0) {
      		result.push(value);
    	}
  }
  return result;
}

function tagsgen(input, email, uname)
{
	return input
		.replace(new RegExp("##email##", "g"), email)
		.replace(new RegExp("##username##", "g"), uname)
		.replace(new RegExp(/##blank_(.*?)##/, "g"), (_, n) => randomstring("blank",n))
		.replace(new RegExp(/##number_(.*?)##/, "g"), (_, n) => randomstring("number",n))
		.replace(new RegExp(/##letter_(.*?)##/, "g"), (_, n) => randomstring("letter",n))
		.replace(new RegExp(/##letterup_(.*?)##/, "g"), (_, n) => randomstring("letterup",n))
		.replace(new RegExp(/##letterlow_(.*?)##/, "g"), (_, n) => randomstring("letterlow",n))
		.replace(new RegExp(/##letternumber_(.*?)##/, "g"), (_, n) => randomstring("letternumber",n))
		.replace(new RegExp(/##letternumberup_(.*?)##/, "g"), (_, n) => randomstring("letternumberup",n))
		.replace(new RegExp(/##letternumberlow_(.*?)##/, "g"), (_, n) => randomstring("letternumberlow",n))
		.replace(new RegExp(/##32bit_(.*?)##/, "g"), (_, n) => randomstring("32bit",n))
}

function onlyUnique(value, index, self) {
  return self.indexOf(value) === index;
}

async function sending(email, uname, no, listTotal)
{
	var smtps = smtplist[no % smtplist.length];
	var link = links[no % links.length];
	var currentlink = links[no % links.length];
	var letter = fs.readFileSync(config.send.LETTER.Folder+"/"+letters[no % letters.length], "utf-8");
	var currentletter = letters[no % letters.length];

	if (config.send.ATTACHMENT.Folder) {
		var attachmentfile = attachments[no % attachments.length];
		var currentattachment = attachments[no % attachments.length];
	}

	link = tagsgen(customrandom(link), email, uname);
	letter = tagsgen(customrandom(letter), email, uname);
	letter = letter.replace(new RegExp("##link##", "g"), link);

	var smtp = {
		host: config.send.SMTP.Host,
		port: config.send.SMTP.Port,
		secure: false,
		requireTLS: true,
		secureConnection: true,
		debug: false,
		logger: false,
		name: tagsgen(customrandom(config.send.SMTP.Hostname), email, uname),
		auth: {
			user: smtps.split(",")[0],
			pass: smtps.split(",")[1]
		}
	};

	const transporter = nodemailer.createTransport(smtp);
	const message = [];

	if (config.send.SEND.Priority) {
		message.priority = customrandom(config.send.SEND.Priority);
	}

	if (config.send.SEND.Returnpath) {
		var returnPath 	 = tagsgen(customrandom(config.send.SEND.Returnpath), email, uname);
		returnPath	   	 = returnPath.replace(new RegExp("##domain##", "g"), smtps.split(",")[0].split("@")[1]);
		message.envelope = {
			"from"  : tagsgen(returnPath, email, uname),
			"to"	: email
		};
	}

	if (config.send.SEND.Messageid) {
		message.messageId = tagsgen(customrandom(config.send.SEND.Messageid), email, uname);
	}

	if (config.send.SEND.Bccmode) {
		message.bcc = email;
		if (config.send.SEND.Bccto) {
			message.to = tagsgen(customrandom(config.send.SEND.Bccto), email, uname);
		}
	} else {
		message.to = email;
	}

	if (config.send.SEND.Frommail) {
		var fromMail = tagsgen(customrandom(config.send.SEND.Frommail), email, uname);
		fromMail	 = fromMail.replace(new RegExp("##domain##", "g"), smtps.split(",")[0].split("@")[1]);
		message.from = tagsgen(customrandom(config.send.SEND.Fromname), email, uname) + " <" + fromMail + ">";
	} else {
		message.from = tagsgen(customrandom(config.send.SEND.Fromname), email, uname) + " <" + smtps.split(",")[0] + ">";
	}

	message.subject = tagsgen(customrandom(config.send.SEND.Subject), email, uname);

	if (config.send.SEND.Replyto) {
		message.replyTo = tagsgen(customrandom(config.send.SEND.Replyto), email, uname);
	}

	if (config.send.HEADER.Useheader) {
		var headers = {};
		var headerObj = config.send.HEADER.Headerval;
		Object.keys(headerObj).forEach(key => {
			headers[tagsgen(customrandom(key), email, uname)] = tagsgen(customrandom(headerObj[key]), email, uname);
		});
		message.headers = headers;
	}

	message.attachments = [];
	if (config.send.ATTACHMENT.Folder) {
		var attachmentcontent = {
			filename : tagsgen(customrandom(config.send.ATTACHMENT.Name), email, uname),
			path : config.send.ATTACHMENT.Folder+"/"+attachmentfile,
			encoding : customrandom(config.send.ATTACHMENT.Encoding)
		};
		message.attachments.push(attachmentcontent);
	}

	var imagecid = "";
	if (config.send.IMAGE.File) {
		imagecid = tagsgen("##letternumber_12##", email, uname);
		var imagecontent = {
			filename : tagsgen(customrandom(config.send.IMAGE.Name), email, uname),
			path : config.send.IMAGE.File,
			encoding : customrandom(config.send.IMAGE.Encoding),
        	cid: imagecid
		};
		message.attachments.push(imagecontent);
	}

	if (config.send.LETTER.Type != "html") {
		message.text = letter;
	} else {
		message.html = letter.replace(new RegExp("##image##", "g"), "cid:"+imagecid);
	}

	if (config.send.ALTERNATIVE.File)
	{
		var alt = fs.readFileSync(customrandom(config.send.ALTERNATIVE.File), "utf-8"); 
		message.text = tagsgen(customrandom(alt), email, uname);
	}

	message.encoding = customrandom(config.send.LETTER.Encoding);
	message.textEncoding = customrandom(config.send.ALTERNATIVE.Encoding);

	transporter.sendMail(message, (error, info) => {
	   	if (error) {
	    	if (config.send.LIST.Logfailed) {
				fs.appendFile('Log/failed.txt', email+"\n", err => { if (err) { console.error(err); }});
			}
			console.log(
				colors.brightBlue(" [+]")+
				colors.brightYellow(" ["+(no+1)+"/"+listTotal+"]\n") +
				colors.brightBlue(" [+]") +
				colors.brightMagenta(" From Name   => "+config.send.SEND.Fromname+"\n") +
				colors.brightBlue(" [+]")+
				colors.brightMagenta(" SMTP        => "+smtps.split(",")[0]+"\n") +
				colors.brightBlue(" [+]")+
				colors.brightCyan(" Letter/Link => "+currentletter+" / "+currentlink)
			);

			if (config.send.ATTACHMENT.Folder) {
				console.log(
					colors.brightBlue(" [+]")+
					colors.brightCyan(" Attachment  => "+currentattachment)
				);
			}

			console.log(
				colors.brightBlue(" [+]")+
				colors.brightCyan(" Sent To     => "+uname+" <"+email+">\n") +
				colors.brightBlue(" [+]")+
				colors.brightRed(" Status      => Failed : "+error+"\n")
			);
	   	} else {
			console.log(
				colors.brightBlue(" [+]")+
				colors.brightYellow(" ["+(no+1)+"/"+listTotal+"]\n") +
				colors.brightBlue(" [+]") +
				colors.brightMagenta(" From Name   => "+config.send.SEND.Fromname+"\n") +
				colors.brightBlue(" [+]")+
				colors.brightMagenta(" SMTP        => "+smtps.split(",")[0]+"\n") +
				colors.brightBlue(" [+]")+
				colors.brightCyan(" Letter/Link => "+currentletter+" / "+currentlink)
			);

			if (config.send.ATTACHMENT.Folder) {
				console.log(
					colors.brightBlue(" [+]")+
					colors.brightCyan(" Attachment  => "+currentattachment)
				);
			}

			console.log(
				colors.brightBlue(" [+]")+
				colors.brightCyan(" Sent To     => "+uname+" <"+email+">\n") +
				colors.brightBlue(" [+]")+
				colors.brightGreen(" Status      => Success\n")
			);
	   	}
	});
}


async function start()
{
	if (config.send.EXPERIMENT.Emailtest) {
		if (config.send.EXPERIMENT.Testafter) {
			var emailTest 	= config.send.EXPERIMENT.Emailtest;
			var listf 		= insertAfter(list, emailTest, config.send.EXPERIMENT.Testafter);
		} else {
			var listf = list;
		}
	} else {
		var listf = list;
	}

	if (config.send.LIST.RemoveDup) {
		listf = listf.filter(onlyUnique);
	}

	console.log(colors.brightCyan(" Send to => "+listf.length+" list"));



	var no = 0;
	listf.forEach((raw_list, n) => {
		setTimeout(() => {
			sending(raw_list.split(",")[1], raw_list.split(",")[0], no, listf.length);
			no++;
		}, n * config.send.SEND.Delay)
	});
}


start();


