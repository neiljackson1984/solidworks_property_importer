<?php

	/*
	This script reads several files and encodes the contents of those files into a big block vba code, with one big string constant for each file.
	
	The output vba code is dumped into a text file, intended for copying and pasting into the vba editor.
	
	We will read two particular javascript files from the current directory, runs them through jsmin, then escapes the result as appropriate to paste into the VBA editor.

	We will also read some bitmap image files and base64 encode them. (it now occurs to me that I could base64 encode everything - both javascript and picture files).
		
	The purpose is to embed the contents of these files as string constants within a vba macro for solidowrks.
	*/


	$outputStream = fopen("javascriptModuleForVba.txt", 'w');
	fwrite($outputStream, encodeForVba(shell_exec("type json2.js | jsmin"), "json2_code"));
	//$output .= encodeForVba(shell_exec("type persistentJsObject.js | jsmin"), "persistentJsObject_code");
	//$output .= encodeForVba(file_get_contents("persistentJsObject.js"), "persistentJsObject_code");
	//fwrite($outputStream, encodeForVba(shell_exec("type persistentJsObject.js"), "persistentJsObject_code"));
	fwrite($outputStream, encodeForVba(shell_exec("type persistentJsObject.js | jsmin"), "persistentJsObject_code"));

	fwrite($outputStream, encodeForVba( base64_encode(file_get_contents("macroFeatureIcon_highlighted.bmp")), "macroFeatureIcon_highlighted_bmp"));
	fwrite($outputStream, encodeForVba( base64_encode(file_get_contents("macroFeatureIcon_regular.bmp")), "macroFeatureIcon_regular_bmp"));
	fwrite($outputStream, encodeForVba( base64_encode(file_get_contents("macroFeatureIcon_suppressed.bmp")), "macroFeatureIcon_suppressed_bmp"));
	
	echo strlen(base64_encode(file_get_contents("macroFeatureIcon_highlighted.bmp")));
	
	fclose($outputStream);


	function encodeForVba($x, $name)
	{
		$maxAllowedLineLength = 1000; //actual limit is 1024 in solidworks vba editor -- I am being generous.
		$chunkLength = $maxAllowedLineLength - 20; 
		
		$escapedString = $x;
		//$escapedString = stripCStyleComments($escapedString);
		$escapedString = str_replace("\t", " ", $escapedString);  //HACK to deal with windows style line endings.
		$escapedString = str_replace("\n", " ", $escapedString); //replace newline with space (This is a hack and will break in cases where string literals in the javascript contain newlines), but for the immediate purpose, it is good enough.
		
		
		$escapedString = str_replace("\"", "\"\"", $escapedString); //replace each quote with two quotes.
		
		$chunks = str_split($escapedString, $chunkLength);
		
		$returnValue = "Public Const $name = \"\"";
		
		foreach($chunks as $chunk)
		{
			//echo "strlen(\$chunk): " . strlen($chunk) . "\n";
			$returnValue .= " _" . "\n";
			$returnValue .= "     " . "& " . "\"" . $chunk . "\"";
		}	
		$returnValue .= "\n\n\n";
		return $returnValue;
			
	}

	
	
	/*
		Strips c-style comments (i.e. comments of the form  //...(EOL)     OR      /* ... * /) from the string argument.
	 * 
	 */
	function stripCStyleComments($x)
	{
		$y = $x;
		$y = preg_replace('!/\*.*?\*/!s', '', $y);  //strips block comments of the form /*...*/
		$y = preg_replace('/\/\/.*$/m', "", $y); //strips single line comments of the form //...
		return $y;
	};
?>