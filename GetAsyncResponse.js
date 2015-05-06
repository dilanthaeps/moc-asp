// JScript File

function AsyncResponse(_surl,func,resXML) {
    this.setUrl(_surl);
    this.setReturn = func;
    
    //Exposed Method
    this.getResponse = function () {
        var req ;
        var f=this.setReturn;
        var resXML_ = resXML;
        //f("Loading ...");
        if (window.XMLHttpRequest){
            req = new XMLHttpRequest();
            if (req.overrideMimeType) {
                     req.overrideMimeType('text/xml');
            }      
            if (req){
                req.onreadystatechange = function(){
                                        if (req.readyState == 4) {
                                            if (req.status == 200 || req.status == 304) {
                                                if (resXML_)
                                                    f(req.responseXML);
                                                else    
                                                    f(req.responseText);
                                            }
                                        }
                                    };
                req.open("GET", this.url, true);
                req.send(null);
            }
        }else if (window.ActiveXObject){
            req = new ActiveXObject("Microsoft.XMLHTTP");
            if (req){
                req.onreadystatechange = function(){
                                        if (req.readyState == 4) {
                                            if (req.status == 200 || req.status == 304) {
                                                f( req.responseText);
                                            }
                                        }
                                    };                     
                req.open("GET", this.url, true);
                req.send();
            }
        } 
    }; 
    // getResponse ends
}
// class ends 



//property setUrl
AsyncResponse.prototype.setUrl = function (surl_) {
    this.url = surl_ ;
    return this;
};
//property getUrl 
AsyncResponse.prototype.getUrl = function () {
    return this.url;
};
