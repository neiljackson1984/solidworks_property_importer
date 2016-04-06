<?php

	/*
	This script reads two particular javascript files from the current directory, runs them through jsmin, then escapes the result as appropriate to paste into the VBA editor.

	This is to embed this javascript code as string constants within a vba macro for solidowrks.
	*/


	$outputStream = fopen("javascriptModuleForVba.txt", 'w');
	fwrite($outputStream, encodeForVba(shell_exec("type json2.js | jsmin"), "json2_code"));
	//$output .= encodeForVba(shell_exec("type persistentJsObject.js | jsmin"), "persistentJsObject_code");
	//$output .= encodeForVba(file_get_contents("persistentJsObject.js"), "persistentJsObject_code");
	//fwrite($outputStream, encodeForVba(shell_exec("type persistentJsObject.js"), "persistentJsObject_code"));
	fwrite($outputStream, encodeForVba(shell_exec("type persistentJsObject.js | jsmin"), "persistentJsObject_code"));

	fclose($outputStream);


	function encodeForVba($x, $name)
	{
		$maxAllowedLineLength = 1000; //actual limit is 1024 in solidworks vba editor -- I am being generous.
		$chunkLength = $maxAllowedLineLength - 20; 
		
		$escapedString = $x;
		$escapedString = stripCStyleComments($escapedString);
		$escapedString = str_replace("\t", " ", $escapedString);  //HACK to deal with windows style line endings.
		$escapedString = str_replace("\n", " ", $escapedString); //replace newline with space (This is a hack and will break in cases where string literals in the javascript contain newlines), but for the immediate purpose, it is good enough.
		
		
		$escapedString = str_replace("\"", "\"\"", $escapedString); //replace each quote with two quotes.
		
		$chunks = str_split($escapedString, $chunkLength);
		
		$returnValue = "Public Const $name = \"\"";
		
		foreach($chunks as $chunk)
		{
			echo "strlen(\$chunk): " . strlen($chunk) . "\n";
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