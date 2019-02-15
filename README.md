Many applications, both Windows and Web, need to connect to the Internet to send or receive data. ASNA Visual RPG Classic apps do not intrinsically have the ability to make HTTP requests. This can be a crippling drawback when you have a legacy enterprise app for which you now have the requirement to send and/or receive data from the Internet. There are still a few third-party COM controls around that can enable AVR to make HTTP requests, but these controls are often troublesome and customers that some no longer work.

This article shows how to extend AVR Classic with AVR for .NET to enable the AVR Classic app to fetch data Json data with an HTTP request. We'll first a look at the AVR for .NET class library project responsible for making the HTTP request to fetch the Json then we'll take a look at the AVR Classic app that consumes that .NET class library. This article provides a simple example of fetching Json. Things like user authentication and product-worthy error handling are omitted--but both can be added. 

This article is the third in a series about integrating AVR for .NET with AVR Classic. The other two are: 

* [How to make .NET's command line avaiable inside Visual Studio](https://asna.com/us/tech/kb/doc/dot-net-command-line)
* [How to extend AVR Classic with AVR for .NET](https://asna.com/us/tech/kb/doc/extend-avr-classic)

Other articles that may be helpful are:

* [How to read and write Json with AVR for .NET](https://asna.com/us/tech/kb/doc/read-write-json)
* [AVR for .NET arrays: dynamically-sized arrays](https://asna.com/us/tech/kb/series/avr-rpg-arrays/dynamic-arrays)

Let's start with a preview of the results. The image below shows an AVR Classic app with a simple subfile. This subfile has been populated with Json data read from the Internet--with a little help from AVR for .NET. You may not need to populate a subfile with Json data, but the subfile is a good way to show results. 

![](https://asna.com/filebin/marketing/article-figures/avr-classic-json/avr-classic-app.png)

### The Json data

First we need some Json test data. The [JSONPlaceholder](https://jsonplaceholder.typicode.com/) site provides several different Json test documents. This article uses its [`users` Json document](https://jsonplaceholder.typicode.com/users). This URL provides ficticious Json data about 10 users and a fragment of it is shown below in Figure 1. 

![](https://asna.com/filebin/marketing/article-figures/avr-classic-json/json-data.png)

<small>Figure 1. Sample Json data from [https://jsonplaceholder.typicode.com/users](https://jsonplaceholder.typicode.com/users.)</small>

The red box in Figure 1 above outlines a single user element in the Json document. Each user in this Json document has nested data in its `address`, `geo`, and `company` values. We'll need a set of AVR for .NET classes to model this data so we can deserialize the Json data into a data format easily accessible by the AVR Classic app. 

### The AVR for .NET class library

To integrate AVR for .NET with AVR Classic, we'll start by building an AVR for .NET class libary. This library will offer AVR Clasic what it sees as a custom COM component with ability to make an HTTP request. 

The AVR for .NET classes that model this data are shown below in Figure 2a. It's generally a best practice to put each class in its own source file, but to mimimize the chunks presented here I'm cheating and putting these four classes in a single source member. 

    Using System.Runtime.InteropServices
    
    DclNameSpace AVRClassicHelper.Http 
    
    BegClass User Access(*Public) Attributes(ComVisible(*True), +
                               ClassInterface(ClassInterfaceType.AutoDual))
        DclProp id Type(*String) Access(*Public) 
        DclProp name Type(*String) Access(*Public) 
        DclProp username Type(*String) Access(*Public) 
        DclProp email Type(*String) Access(*Public) 
        DclProp address Type(AddressInfo) Access(*Public) 
        DclProp phone Type(*String) Access(*Public) 
        DclProp website Type(*String) Access(*Public) 
        DclProp company Type(CompanyInfo) Access(*Public) 
    EndClass
    
    BegClass AddressInfo Access(*Public) Attributes(ComVisible(*True), +
                                 ClassInterface(ClassInterfaceType.AutoDual))
        DclProp street Type(*String) Access(*Public) 
        DclProp suite Type(*String) Access(*Public) 
        DclProp city Type(*String) Access(*Public) 
        DclProp zipcode Type(*String) Access(*Public) 
        DclProp geo Type(GeoInfo) Access(*Public) 
    EndClass
    
    BegClass GeoInfo Access(*Public) Attributes(ComVisible(*True), +
                             ClassInterface(ClassInterfaceType.AutoDual))
        DclProp lat Type(*String) Access(*Public) 
        DclProp lng Type(*String) Access(*Public) 
    EndClass
    
    BegClass CompanyInfo Access(*Public) Attributes(ComVisible(*True), + 
                                 ClassInterface(ClassInterfaceType.AutoDual))
        DclProp name Type(*String) Access(*Public) 
        DclProp catchphrase Type(*String) Access(*Public) 
        DclProp bs Type(*String) Access(*Public) 
    EndClass

<small>Figure 2a. Sample Json data from https://jsonplaceholder.typicode.com/users.</small>

Figure 2a is a single source member that provides four classes: 

* User
* AddressInfo
* GeoInfo
* CompanyInfo 

Notice how the structure of these four classes echoes exactly the nested structure presented by the Json test data. Ultimately, we'll be able to fetch a user property with a nested object syntax like this:

    DclFld City Type(*String)
    
    City = User.Address.City
    
It's very important that the structures created to represent the Json data do so accurately. Take your time and declare your data description classes very carefully. The ability to deserialize the incoming Json into a .NET object depends on their schema correctly echoing the Json schema.

You'll notice that the .NET attributes `ComVisible` and `ClassInterface` have been applied to all four of the data description classes in Figure 2a. These attributes are necessary to surface .NET classes (and their properies and members) to COM. [This asna.com article goes into more detail about these attributes.](https://asna.com/us/tech/kb/doc/extend-avr-classic) These attributes are applied to all of four of the classes in Figure 2a.

> Assigning the `ComVisible` and `ClassInterface` attributes used to expose .NET classes to COM don't affect the ability of those classes to also be consumed by .NET. These classes can still be used in .NET-only projects. There may be some performance penalty imposed so watch for that--but that doesn't appear to a significant issue. 

The second AVR for .NET class needed is one to make the HTTP request to fetch the Json. That `Request` class is shown below in Figure 2b. 

    Using System
    Using System.IO
    Using System.Text 
    Using System.Web
    Using System.Net 
    Using NewtonSoft.Json
    
    DclNameSpace AVRClassicHelper.Http 
    
    BegClass Request Access(*Public)
        DclProp HttpStatus Type(*Integer4) Access(*Public) 
        DclProp ErrorMessage Type(*String) Access(*Public) 
    
        BegFunc GetRequest Type(*String) Access(*Public) 
            DclSrParm Url Type(*String) 
            
            DclFld encoding Type(ASCIIEncoding) New()
            DclFld req Type(HttpWebRequest) 
            DclFld res Type(HttpWebResponse) 
            DclFld responseStream Type(Stream) 
            DclFld responseString Type(*String) 
            DclFld sr Type(StreamReader) 
           
            req = WebRequest.Create(Url) *As HttpWebRequest         
            req.Method = "GET"
    
            Try
                res = req.GetResponse() *As HttpWebResponse
                *This.HttpStatus = res.StatusCode
                *This.ErrorMessage = String.Empty
    
            Catch ex1 Type(WebException)
                If ex1.Status <> WebExceptionStatus.Success
                    res = ex1.Response *As HttpWebResponse 
                    *This.HttpStatus = res.StatusCode
                    *This.ErrorMessage = ex1.Message 
                    LeaveSr *Nothing 
                EndIf 
    
            Catch ex2 Type(Exception) 
                *This.HttpStatus = 0
                *This.ErrorMessage = ex2.Message 
                LeaveSr *Nothing 
            EndTry 
    
            If (res.StatusCode = HttpStatusCode.OK)
                responseStream = res.GetResponseStream()
                sr = *New StreamReader(responseStream) 
                responseString = sr.ReadToEnd() 
                sr.Close()
            Else 
                *This.HttpStatus = 0
                *This.ErrorMessage = res.StatusDescription
                LeaveSr *Nothing 
            EndIf
            
            LeaveSr responseString 
        EndFunc                 
    
        BegFunc GetJson Access(*Public) Type(User) Rank(1)
            DclSrParm Url Type(String) 
            DclFld JsonString Type(*String) 
            
            DclArray UserList Type(User) Rank(1) 
    
            JsonString = GetRequest(Url)         
            If *This.HttpStatus = 200 
                UserList = JsonConvert.DeserializeObject(JsonString, +
                                    *TypeOf(User[])) *As User[]
                LeaveSr UserList 
            Else 
                LeaveSr *Nothing 
            EndIf 
        EndFunc 
    
    EndClass

<small>Figure 2b. The AVR for .NET Request class.</small>

The `Request` class has two properties:

* HttpStatus - This property reports the HTTP status of the most recent request. If the request succeeded this value will be 200, otherwise an error occured.
* ErroMessage - This field reports the error message when an error occurs.

The `Request` class has two methods:

* `GetRequest` - This method uses the URL passed to it to make an HTTP Get request. If the request succeeds this method returns the string value of the response. When used to fetch Json, this will be a string value of the Json. A short sidebar at the end of this article goes into a little more detail on the `GetRequest` method.
* `GetJson` - This method is wrapper around the more general-purpose `GetRequest` method to make a Json request. It returns a .NET object that is deserialized from the Json string returned. In this example, an array of the User class (from Figure 2a) is returned. [This aricle explains the Json deserialization process in detail.](https://asna.com/us/articles/newsletter/2016/q3/read-write-json)

Note that nothing in the `Request` class is surfaced directly to COM. 

The third and final AVR for .NET class required for this example is a class to surface Figure 2b's GetJson method and some other necessary properties. My convention is to call this class `ComBridge` and it is shown in Figure 2c below. 

> Debugging between AVR Classic and AVR for .NET is challenging. You can't interactively debug from one environment to the other. The ComBridge class makes it easy to package and test exactly what the COM app needs--and this minimizes the .NET debugging required. As you build .NET components for AVR Classic consumption code defensively and test your .NET components carefully before attempting to consume them with COM. A .NET test harness to test your .NET work first is a good way to avoid debugging pain later. 

    Using System
    Using System.Text
    Using System.Runtime.InteropServices
    
    DclNameSpace AVRClassicHelper.Http 
    
    BegClass ComBridge Access(*Public)  +
                       Attributes(ComVisible(*True), +
                       ClassInterface(ClassInterfaceType.AutoDual))
    
        DclProp HTTPStatus Type(*Integer4) Access(*Public) Attributes(ComVisible(*True))
        DclProp ErrorMessage Type(*String) Access(*Public) Attributes(ComVisible(*True))
        DclArray Users Type(User) Rank(1) Access(*Public)  
        
        BegFunc CallGet Access(*Public) Attributes(ComVisible(*True)) Type(*Integer4)  
            DclSrParm Url Type(*String) 
    
            DclFld Req Type(AVRClassicHelper.Http.Request) New() 
            Users = Req.GetJson(Url)       
            
            *This.HTTPStatus = Req.HTTPStatus
            *This.ErrorMessage = Req.ErrorMessage
    
            If Req.HTTPStatus = 200 
                LeaveSr Users.Length
            EndIf 
    
            Users = *Nothing 
            LeaveSr -1 
        EndFunc 
    
        BegFunc GetUser Access(*Public) Type(User) Attributes(ComVisible(*True))
            DclSrParm Index Type(*Integer4) 
    
            If Users = *Nothing 
                LeaveSr *Nothing 
            EndIf 
    
            If Index > Users.Length - 1 
                LeaveSr *Nothing
            Else 
                LeaveSr Users[Index]
            EndIf 
        EndFunc  
    
    EndClass

<small>Figure 2c. The ComBridge class which surfaces .NET functionality to AVR Classic.</small>

Like the data classes in Figure 2a, the `ComBridge` class is also decorated with the `ComVisible` and `ClassInterface` attributes. 

`ComBridge` has three properties:

* HttpStatus - This property exposes the Request's class's HttpStatus property.
* ErrorMessage  - This property exposes the Request's class's HttpError message property.
* Users - This property is an array of users, which is this is example is populated with Json. When you fetch a Json array, you rarely know how many array elements will be returned (the test data we're using is hardcoded to 10 elements, but that hardcoding rarely happens in the real world). The Users array is a [Ranked array](https://asna.com/us/tech/kb/series/avr-rpg-arrays/dynamic-arrays); that's an array type that AVR Classic doesn't support; AVR can't use this object directly. To avoid it being surfaced to COM (where it would cause a runtime error), it is marked `ComVisible(*False)`. 

> Limit the data types you surface to COM from .NET to scalar types (strings and numbers, essentially) and data classes (like the four we're using from Figure 2a). Even simple .NET objects like date data types can cause issues with COM. Use getter functions (like this example does) and keep the interface between the two environments simple. 

`ComBridge` surfaces two methods to AVR Classic: 

* CallGet - This method uses the Request's class's GetJson method to deserialize the Json to populate the `Users` ranked array. Because AVR Classic can't directly access the `Users` array, this method returns to AVR Classic the number of elements read. 
* GetUser - This method surfaces a given element of the `User` array. You'll see in a moment that AVR Classic uses this method in a loop to fetch each user.

After compiling the .NET project, .NET's `RegAsm `utility must be used to create COM-based type library needed for AVR Classic. [`RegAsm` and how to use it is explained in this article.](https://asna.com/us/tech/kb/doc/dot-net-command-line) After running `RegAsm` you'll see that library in the same folder as the .NET DLL--the only difference is the COM library has a `.tlb` extension. `RegAsm` created the library and registered with COM on your system. 

### The AVR classic app to consume the .NET class library

Consuming the AVR for .NET project's DLL with AVR Classic is pretty simple. With a new project started, we first need to set a reference to the COM class library the AVR for.NET project created. In this example, that DLL is named AVRClassicHelper_Http. The .NET project name was AVR ClassicHelper.HTTP and when `RegAsm` compiled the COM type library it swaps out the period for an underscore. Figure 3a below shows AVR Classic's References window with this reference set. 

![](https://asna.com/filebin/marketing/article-figures/avr-classic-json/references.png)

<small>Figure 3a. AVR Classic's References window</small>

Having set that reference, we can use AVR Classic's Object Browser to see what that reference makes available to AVR Classic, as shown below in Figure 3b. 

![](https://asna.com/filebin/marketing/article-figures/avr-classic-json/classic-object-browser.png)

<small>Figure 3b. AVR Classic's Object Browser window</small>

AVR Classic's Object Browser shows what .NET components the reference made available: the four data classes (`User`, `AddressInfo`, `CompanyInfo`, and `GeoInfo`) and the `ComBridge class`. The Object Browser view in Figure 3b shows methods and properties that the `ComBridge` makes available. You'll use AVR Classic's Object Browser frequently to ensure the members and properties you think should be there are actually there and to see how to declare the classes (in the bottom of the Object Browser window).

> You also notice that the `ComBridge` class surfaces members (the properties `Equals`, `GetType`, `GetHashCode`, `GetType`, and the method `ToString`) that weren't explicitly defined in the `ComBridge` class in the .NET project. These members are aritifacts of .NET object inheritance and you can generally ignore them.

The AVR Classic code to use the `ComBridge` class is shown below in Figure 4a:. 
    
    DCLFLD httpGetJson TYPE(AVRClassicHelper_Http.ComBridge)  
    DCLFLD User TYPE(AVRClassicHelper_Http.User) 
    
    labelResult.Caption = ''
    
    BEGSR CommandButton1 Click
        DclFld UserCount TYPE(*Integer) Len(4) 
        DclFld Url Type(*String) 
        DclFld i Type(*Integer) Len(4) 
    
        SetMousePtr *HourGlass
        
        Url = 'https://jsonplaceholder.typicode.com/users'
        UserCount = httpGetJson.CallGet(Url) 
        
        If httpGetJson.HttpStatus = 200    
            SetMousePtr *Dft 
            If UserCount < 0
                MsgBox 'Error reading Json data'
                LeaveSr 
            EndIf 
    
            labelResult.Caption = 'Json rows read: ' + %TRIM(%CHAR(%EDITC(UserCount, 'J'))) 
    
            subfileUsers_RRN = 0 
            subfileUsers.ClearObj()
            Do FromVal(0) ToVal(UserCount-1) Index(i) 
                User = httpGetJson.GetUser(i)
                WriteSubfileRow()
            EndDo 
        Else
            MsgBox Msg(httpGetJson.ErrorMessage) 
        EndIf        
    ENDSR
    
    BegSr WriteSubfileRow
        subfileUsers_RRN = subfileUsers_RRN + 1
        Id = User.Id
        Name = User.Name
        Email = User.Email
        City = User.Address.City
        ZipCode = User.Address.ZipCode
        Company = User.Company.Name
    
        Write subfileUsers     
    EndSr
    
<small>Figure 4a. The AVR Classic code to consume the .NET components.</small>    

This line from the code above calls the .NET `CallGet()` method, passing it a URL: 

    UserCount = httpGetJson.CallGet(Url) 

If the call is successful, `UserCount` indicates the number of Json elements available and `httpGetJson.HttpStatus` will be 200. If the call isn't successful `UserCount` is -1 and the `httpGetJson.HttpStatus` code is the HTTP status code received from the HTTP request. If an error occurred the error message is in the `httpGetJson.ErrorMessage` property.

Because AVR Classic can't access the .NET array of users directly, it uses a loop and the .NET `httpGetJson.GetUser` function to fetch each user. When a user is fetched `WriteSubFile` is called to add that user to the subfile. 

The results of the code in Figure 4a are shown below in Figure 4b:

![](https://asna.com/filebin/marketing/article-figures/avr-classic-json/avr-classic-app.png)


### Summary 

Integrating AVR Classic with AVR for .NET is an intermediate/advanced topic to be sure. However, once you master the basics, it's actually pretty easy to do and the power and possiblities that .NET can provide to COM are nearly endless. 

Considerations:

* The clients need the same version of AVR for .NET's runtime installed.
* The clients also need the .NET Framework installed but for any Windows 7/8/10 box it will already be there. However, it might be an old version so you might need to [update the version of the .NET Framekwork installed.](https://www.microsoft.com/en-us/download/details.aspx?id=55170) For best results, make your client PC's .NET Framework version match your development version.
* Copy the AVR for .NET DLL the projct produces to each client PC. 
* Register the DLL with the `RegAsm` utility (which the .NET Framework provides) [as explained here.](https://asna.com/us/tech/kb/doc/extend-avr-classic)
* Test your .NET code before trying to integrate with AVR Classic. Debugging between the two environments is frustrating. 
* Limit the data you pass from .NET to AVR Classic to core scalar values and simple data classes. 
* Don't start with your biggest, most important project! Start simple and build from there.

<hr>

### Sidebar: A quick note on fetching the Json respose

The AVR for .NET `GetRequest` method in Figure 2b above has about 45 lines of code, however much of that code is error handling. The core facility to issue an HTTP request and convert its response to a string is shown below in Figure A. 

    DclFld req             Type(HttpWebRequest) 
    DclFld res             Type(HttpWebResponse) 
    DclFld responseStream  Type(Stream) 
    DclFld responseString  Type(*String)            
    DclFld sr              Type(StreamReader) 
    
    // Make an HTTP GET request return its response object to the res variable..
    req = WebRequest.Create(Url) *As HttpWebRequest         
    req.Method = "GET"
    res = req.GetResponse() *As HttpWebResponse
    
    // Convert the response into a string. 
    responseStream = res.GetResponseStream()
    sr = *New StreamReader(responseStream) 
    responseString = sr.ReadToEnd() 
    sr.Close()            

<small>Figure A. Traditional AVR for .NET code to work with HTTP.</small>

The code above uses the .NET [`HttpWebRequest`](https://docs.microsoft.com/en-us/dotnet/api/system.net.httpwebrequest?view=netframework-4.7.1) and [`HttpWebResponses`](https://docs.microsoft.com/en-us/dotnet/api/system.net.httpwebresponse?view=netframework-4.7.1) classes from the `System.Net` namespace to issue an HTTP GET request and process its response. 

We've been doing HTTP work with AVR for .NET for a long time and have always used the `HttpWebRequest/HttpWebResponse` APIs. The grungy part of using `HttpWebRequest/HttpWebResponse` isn't issuing the HTTP request, it's the four mysterious lines of code required to convert the response into a string. It occurred to me during this .NET->COM project that maybe there are better ways now to work with HTTP in .NET.

[This article](https://code-maze.com/different-ways-consume-restful-api-csharp/) lead me to the RestSharp open source project. With nearly [17m downloads on nuget.org](https://www.nuget.org/packages/RestSharp/) RestSharp must have something going for it. I gave it a quick spin with AVR for .NET and was impressed with its concise API and direct way of doing things. With RestSharp the code in Figure A above is reduced to the code in Figure B below:

    DlFld Client Type(RestSharp.RestClient) 
    DclFld Response Type(RestSharp.IRestResponse) 
    DclFld ResponseString Type(*String) 

    Client = *New RestSharp.RestClient(url)
    Response = Client.Execute(*New RestSharp.RestRequest())
    ResponseString = Response.Content 
            
<small>Figure B. The RestSharp equivalent of Figure A.</small>

The RestSharp API is comprehensive and features for authentication and serialization baked in. It looks like a promising API for .NET HTTP work.