// Prompt the user for the URL
WScript.Echo("Enter the URL you want to parse:");

// Set the URL
var url = WScript.StdIn.ReadLine();

if (!url) {
    WScript.Echo("No URL provided. Exiting...");
    WScript.Quit();
}

// Create an HTTP request object
var http = new ActiveXObject("MSXML2.XMLHTTP");

try {
    http.open("GET", url, false); // Synchronous request
    http.send();

    if (http.status == 200) {
        // Parse the HTML content
        var htmlContent = http.responseText;

        // Extract data
        var h1Match = htmlContent.match(/<h1[^>]*class="[^"]*h2[^"]*"[^>]*>(.*?)<\/h1>/);
        var title = h1Match ? h1Match[1] : "No Title Found";

        var metaMatch = htmlContent.match(/<meta[^>]*name="description"[^>]*content="([^"]*)"[^>]*>/);
        var description = metaMatch ? metaMatch[1] : "No Description Found";

        var divMatch = htmlContent.match(/<div[^>]*class="[^"]*image[^"]*"[^>]*>/);
        var backgroundImage = null;
        if (divMatch) {
            var styleMatch = divMatch[0].match(/style="([^"]*)"/);
            if (styleMatch) {
                var backgroundImageMatch = styleMatch[1].match(/background-image:\s*url\(["']?([^"')]+)["']?\)/);
                backgroundImage = backgroundImageMatch ? backgroundImageMatch[1] : null;
            }
        }

        // Generate local HTML file
        var fso = new ActiveXObject("Scripting.FileSystemObject");
        
	var firstWord = title.split(" ")[0];
	var fileName = makeSafeFileName(firstWord) + ".html";	
	    
        var file = fso.CreateTextFile(fileName, true);

        // Write HTML content
        file.WriteLine("<!DOCTYPE html>");
        file.WriteLine("<html lang='en'>");
        file.WriteLine("<head>");
        file.WriteLine("<meta charset='UTF-8'>");
        file.WriteLine("<meta name='viewport' content='width=device-width, initial-scale=1.0'>");
        file.WriteLine("<title>" + title + "</title>");
        file.WriteLine("<style>");
        file.WriteLine("body { font-family: Arial, sans-serif; margin: 20px; }");
        file.WriteLine(".content { margin-bottom: 20px; }");
        file.WriteLine(".image { max-width: 100%; height: auto; }");
        file.WriteLine(".button { margin-left: 10px; padding: 5px 10px; cursor: pointer; }");
        file.WriteLine("</style>");
        file.WriteLine("<script>");
        file.WriteLine("function copyToClipboard(text) { navigator.clipboard.writeText(text); alert('Copied: ' + text); }");
        file.WriteLine("</script>");
        file.WriteLine("</head>");
        file.WriteLine("<body>");
        file.WriteLine("<h1 class='content'>" + title + "<button class='button' onclick=\"copyToClipboard('" + title + "')\">Copy</button></h1>");
        file.WriteLine("<p class='content'>" + description + "<button class='button' onclick=\"copyToClipboard('" + description + "')\">Copy</button></p>");
        if (backgroundImage) {
            file.WriteLine("<img src='" + backgroundImage + "' alt='Background Image' class='image'>");
        } else {
            file.WriteLine("<p>No image found.</p>");
        }
        file.WriteLine("</body>");
        file.WriteLine("</html>");

        file.Close();
        WScript.Echo("HTML file generated: " + fileName);
		
		// Open the generated HTML file in the default web browser
		var shell = new ActiveXObject("WScript.Shell");
		shell.Run(fileName);
		
    } else {
        WScript.Echo("Failed to fetch the webpage. Status: " + http.status);
    }
} catch (error) {
    WScript.Echo("An error occurred: " + error.message);
}

// Function to make a string safe for filenames
function makeSafeFileName(name) {
    return name.replace(/[^a-zA-Z0-9_-]/g, "_"); // Replace all invalid characters with "_"
}
