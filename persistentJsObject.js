//thanks to https://blogs.msdn.microsoft.com/david.wang/2006/07/04/howto-convert-between-jscript-array-and-vb-safe-array/ for these two array conversion functions.
//It doesn't look like these functions support nested arrays.

function VB2JSArray( objVBArray )
{
	//return new VBArray( objVBArray ).toArray();

	var temp = new VBArray( objVBArray ).toArray();
	var values = [];
	for(var key in temp)
	{
		values.push(temp[key]);
	}
	return values;
};

function JS2VBArray( objJSArray )
{
	var dictionary = new ActiveXObject( "Scripting.Dictionary" );
	for ( var i = 0; i < objJSArray.length; i++ )
	{
		dictionary.add( i, objJSArray[ i ] );
	}
	return dictionary.Items();
};


//converts an activex dictionary object into a javascript object.
function dictionaryToObject( dict )
{
	var x = new Object();

	var keys = VB2JSArray(dict.Keys());
	//for( var key in keys)
	for( var i=0; i<keys.length; i++)
	{
		var key = keys[i];
		var value = dict.Item(key);
		if( isActiveXDictionary(value) ){
			x[key] = dictionaryToObject(value);
		} else if (isVBArray(value)) {
			x[key] = VB2JSArray(value);
		} else {
			//we are assuming (without checking as thoroughly as we ought to) that dict.Item(key) is a simple atomic value (a string or a number) rather than an object or an array
			x[key] = value;
			//x[key] = i;
		}
	}
	// x = VB2JSArray(dict.Keys());
	// x = isActiveXDictionary(dict);
	// x=i;
	// x =  dict.Count;
	//x=keys.length;
	return x;
	
};

//converts a javascript object into a (possibly nested) dictionary
function objectToDictionary(x)
{
	var dict = new ActiveXObject("Scripting.Dictionary");
	for (var key in x)
	{
		var value = x[key];
		if( isObject(value)){
			dict.add(key, objectToDictionary(value));
		//} else if (false && Array.isArray(value)) {  //"Array.isArray()"  is apparently not supported by windows scripting, so I am simply not going to use  arrays for now.
		} else if (isArray(value)) {  //"Array.isArray()"  is apparently not supported by windows scripting, so I am simply not going to use  arrays for now.
			dict.add(key, JS2VBArray(value));
		} else {
			dict.add(key, value);
		}
		//dict.add(key, 55);
	}
	return dict;
};

//converts a javascript object into a (non-nested (i.e. flat)) dictionary.  If the object is nested, the keys of the dictionary contain dots as separators.
function objectToFlatDictionary(x, keyPrefix)
{
	//this.dictionaryData; //an associative array/object that will be converted to a dictionary and returned whenever this function finishes with nestingLevel==0.
	if(typeof objectToFlatDictionary.nestingLevel === 'undefined'){objectToFlatDictionary.nestingLevel = 0};
	var keyPrefix;
	if(typeof keyPrefix === 'undefined'){ keyPrefix = "";}
	
	if (objectToFlatDictionary.nestingLevel==0) {objectToFlatDictionary.dictionaryData = {}; }; //initialize dictionaryData to an empty object
	objectToFlatDictionary.nestingLevel++;
	
	for (var key in x)
	{
		var value = x[key];
		if( isObject(value)){
			objectToFlatDictionary(value, keyPrefix + key + ".");
			//objectToFlatDictionary.dictionaryData["" + keyPrefix + key + "OBJECT"] = "<OBJECT>";
		} else if (isArray(value)) { 
			//objectToFlatDictionary.dictionaryData[keyPrefix + key] = JS2VBArray(value);
			//for (var arrayKey in value.keys())
			// for (var arrayKey in value)
			// {
				// objectToFlatDictionary(value[arrayKey], keyPrefix + key + "[" + arrayKey + "]" + ".");
				// objectToFlatDictionary(value[arrayKey], keyPrefix + key + "[" + arrayKey + "]" + ".");
				// objectToFlatDictionary.dictionaryData[keyPrefix + key + "[" + arrayKey + "]"] = "ahoy there";
			// }
			// objectToFlatDictionary.dictionaryData["ahoy"] = "ahoy there";
			// objectToFlatDictionary.dictionaryData["arrayLength"] = value.length;
			// objectToFlatDictionary.dictionaryData["typeofstring"] = typeof("abc");
			// objectToFlatDictionary.dictionaryData["isobject_string"] = isObject("abc");
			// objectToFlatDictionary.dictionaryData["isarray_string"] = isArray("abc");
			
			
			//test to see how this will handle sparse arrays
			// // value = new Array();
			// // value[4] = 123.456;
			// // value[17] = 234.567;
			// // value[-3] = 345.678;
			// // value['foo'] = 456.789;
			//this worked as desired, producing two parameters, exactly as expected.
			//The "-3" key was not heeded unless I changed the initializer of the below FOR loop to start i at less than or equal to -3.
			// (value.length is 18 (i.e. highestNaturalNumberIndex + 1)).  The -3 and 'foo' keys have no effect on the value of value.length.
			
			//now something harder:
			// value = new Array();
			// value["ahoy"] = 123.456;
			// value["there"] = 234.567;
			//this did not cause any errors, but neither of the two parameters was extracted.
			
			
			//we construct the object tempObject, whose keys are the rather strange "[0]", "[1]", "[2]", etc.
			var tempObject = new Object();
			var length;
			length = value.length;
			for(var i = 0; i<length; i++)
			{
				if(i in value) 
				{
					tempObject["[" + i + "]"] = value[i];
				}
			}
			objectToFlatDictionary.dictionaryData[keyPrefix + key + "." + "length"] = value.length; //for convenience, we include the array length as a parameter that will appear amongst the Solidworks custom properties.
			
			//Then we invoke objectToFlatDictionary() upon tempObject almost the same way as in the case where value is an object.  The only difference in this case (that is, when value is an array),
			// is that we omit the terminal "." from the second argument to objectToFlatDictionary.  This will mean that, instead of ending up with, for instance,
			// the parameter names { "a.b.c.[0]", "a.b.c.[1]", ...}, we will end up with the parameter names {"a.b.c[0]", "a.b.c[1]", ...}, which is what we want.
			objectToFlatDictionary(tempObject, keyPrefix + key );
			
		} else {
			objectToFlatDictionary.dictionaryData[keyPrefix + key] = value;
		}
	}
	
	objectToFlatDictionary.nestingLevel--;
	if (objectToFlatDictionary.nestingLevel == 0) {return objectToDictionary(objectToFlatDictionary.dictionaryData);} else {return null;}
};


function isObject(x)
{
	return  (typeof x  === "object") && (x !== null) && !isArray(x);
}

//This is really just a stub: it returns true if we can call an "Items()" method on the argument without exception and false otherwise.  IT does not, as of yet, truly check if the argument is an activex dictionary
function isActiveXDictionary(x)
{

	try{
		x.Items();
	} catch (e) {
		return false;
	}
	return true;
};

//this is also a stub function at the moment.  It does not check if the argument is really a VBArry, but it will serve the purpose for my intended application
function isVBArray(x) 
{
	try{
		x.lbound();
	} catch (e) {
		return false;
	}
	return true;
}


//define a class persistentSettings.
var persistentSettings = function()
{
	this.statusText = "constructor called.";
	//this.arbitraryText = Array.isArray([2,3,4]);
	return 0;
};

persistentSettings.prototype.init = function(urlOfFile)
{
	this.urlOfFile = urlOfFile;
	var filesystem = new ActiveXObject("Scripting.FileSystemObject");
	
	
	//check if file exists.  If it does not, create a new empty object.
	if(filesystem.FileExists(urlOfFile)){
		var inputFile = filesystem.OpenTextFile(this.urlOfFile,1);
		this.settings_object = JSON.parse(inputFile.ReadAll());
		inputFile.Close(); //it is good to close the file right away so other scripts can acccess it
		this.FileExists = true;
	} else { 
		this.settings_object = new Object();
		this.FileExists = false;
		//we are not creating the file here; the file will be created when Commit() is called.
	}
		
	this.dictionary = objectToDictionary(this.settings_object);
	//this.flatDictionary = objectToFlatDictionary(this.settings_object, "");
	this.statusText = "init called.";
};

persistentSettings.prototype.Init = persistentSettings.prototype.init;

//When calling a function from VB, if the function has no arguments,
//VB will tend to evaluate symbol as the sring containing the jscript code
//that defines the function, rather than the result of evaluating the function.
//A workaround is to call every function with an argument, even if only a dummy argument.

//this function writes the persisitent settings to disk.	
persistentSettings.prototype.commit = function(dummy)
{
	var filesystem = new ActiveXObject("Scripting.FileSystemObject");
	this.settings_object = dictionaryToObject(this.dictionary);
	var outputFile = filesystem.OpenTextFile(this.urlOfFile,2 /*file mode: write*/,true /*create new file if necessary: yes*/  );
	
	// this.settings_object.foo = [0,1,2];
	// this.settings_object.bar = VB2JSArray(JS2VBArray([0,1,2]));
	// var parseable = "{\"a\":[1,2,3]}";
	// this.settings_object.parseable = parseable;
	// this.settings_object.originalBaz = this.settings_object.baz;
	// this.settings_object.baz = VB2JSArray(JS2VBArray( JSON.parse("{\"a\":[1.7,2.3,3.4]}").a ));
	// this.settings_object.bazIsArray = isArray(this.settings_object.baz);
	// this.settings_object.origianlBazIsArray = isArray(this.settings_object.originalBaz);
	// this.settings_object.result1 = isVBArray(JS2VBArray([0,1,2,3]));
	// this.settings_object.result2 = isActiveXDictionary(JS2VBArray([0,1,2,3]));
	// this.settings_object.result3 = 
		// dictionaryToObject(
			// objectToDictionary(
				// {a:[0,1,2,3]}
			// )
		// );
	// this.settings_object.result4 = isObject([0,1,2,3]);
	
	
	
	outputFile.Write(
		JSON.stringify( this.settings_object  , null, 3)
	);
	outputFile.Close(); //it is good to close the file right away so other scripts can acccess it
};

persistentSettings.prototype.getJSON = function(dummy)
{
	this.settings_object = dictionaryToObject(this.dictionary);
	return JSON.stringify( this.settings_object  , null, 3);
};

persistentSettings.prototype.Item = function(keyString, newValue)
{
	//the newValue argument (with default value null) lets us use this same item() function
	//to get and set items.
	// It seems that windows scripting host does not support the standard syntax for default function arguments (which would be "newValue=null", above), 
	// however, if we invoke this function from VB and pass only one argument, newValue is simply undefined, which we can test for.
	if((newValue !== null) && (newValue !== undefined))
	{
		return this.setItem(keyString, newValue);
	}
	
	var keys = keyString.split(".");

	var temp = this.settings_object;
	for(var i = 0; i<keys.length; i++)
	{	
		var key = keys[i];
		temp = temp[key];
		if(temp === undefined){return undefined;} //returning undefined will cause the cypressVB function IsEmpty() (operating on the return value) to return true.  this provides a way for us to test for undefined keys within the cypress VB script.
	}

	return temp;
};

persistentSettings.prototype.setItem = function(keyString, value)
{
	this.settings_object = dictionaryToObject(this.dictionary); //we do this just in case any changes have been made to the dictionary that are not yet synced to the object.
	var keys = keyString.split(".");
	var lastKey = keys[keys.length-1];
	var temp = this.settings_object;
	for(var level = 1; level<=keys.length-1; level++)
	{	
		var key = keys[level-1];
		if(!isObject(temp[key])){temp[key] = new Object();} //this lets us add new sub objects implicitly, without the objects having to exist to begin with.  However, beware that if temp[key] was a non-object to begin with, that non-object vlaue will be wiped out.
		temp = temp[key];
	}
	//at this point, temp is the object that we want to modify and lastKey is the key to the value within temp that we want to modify.
	temp[lastKey] = value;
	this.dictionary = objectToDictionary(this.settings_object);

	return value;
};

persistentSettings.prototype.getFlatDictionary = function(dummy)
{
	return objectToFlatDictionary(this.settings_object);
};
 
function isArray(obj) {
    return obj instanceof Array;
};
