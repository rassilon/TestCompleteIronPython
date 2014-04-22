TestComplateIronPython
======================

TestComplete IronPython Sample


As part of an architectural prototyping phase I wrote a couple of IronPython wrapping classes around the C# .Net var/VarDelegate class.

 Totally unsupported, but worth noting for anyone who wants to care.

 These two helper classes let me successfully run the C# Order_JS test script and the Hello mspaint JavaScript tests when translated into Pythonese. (i.e. change true to True, etc...)

 This way the syntax in a .Net environment is more like the syntax in JavaScript and you don't have to worry about TestComplete's problem with dying when an exception hits a JavaScript file sandbox boundary.

 The IronPython debugging experience is where this approach fell apart for me. It seemed like the TestComplete COM interaction timed out on me or something while I was debugging in Visual Studio. But a print/Log.* based debug model worked well enough for me to write these classes.

 The only line of code I had to change from the C# Orders_JS script file was the setting of the .Net property using ".Text" to calling the setter method "set_Text" directly like other parts of the JavaScript file did.

 NOTE: No attempt was made to add __*__ methods for ALL numerical operations. Just the ones necessary to make the ported test scripts happy.
 Fortunately as you can see the pattern is fairly straightforward.

 NOTE2: Setting a value via Python [] notiation is not yet implemented.

 NOTE3: var.VarDelegate's static dlgtsForCall seems manifestly unsafe from a threading perspective. YUCK.

 The two classes involved are:
 * TestComplete: This exposes the TC globals as properties and also exposes the usual RunTest/StopTest methods.
 * WrappedObj: This wraps the .Net "var/var.VarDelegate" classes with a Python friendly syntax.

 There is also a WrappedCollection object that I don't think is necessary any longer since WrappedObj supports all of the necessary operations.

 Bill

