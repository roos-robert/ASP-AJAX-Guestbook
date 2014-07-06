function makeRequest(page,divid)
{
   var http_request = getHttpRequest();

   http_request.onreadystatechange = function() { handle_request(http_request,divid); };
   http_request.open('GET', page, true);
   http_request.send(null);
   
   return false;
}

function handle_request(handle,divid)
{
   if (handle.readyState != 4) {
      document.getElementById(divid).innerHTML = 'Anropet misslyckades, kontakta oss om detta!';
      return;
   }
   
   if (handle.status && handle.status != 200) {
      document.getElementById(divid).innerHTML = 'Ett fel har uppstått. Felkod: ' + handle.status;
      return;
   }
   
   document.getElementById(divid).innerHTML = handle.responseText;
}

   
//Returnerar ett XMLHttpRequest objekt
function getHttpRequest()
{
   var handle = false;

   if (window.XMLHttpRequest) { // Firefox, Opera, Safari
      handle = new XMLHttpRequest();

   } else if (window.ActiveXObject) { // Internet Explorer
      try {
         handle = new ActiveXObject("Msxml2.XMLHTTP");
      } catch (e) {
         try {
            handle = new ActiveXObject("Microsoft.XMLHTTP");
         } catch (e) {}
      }
   }

   if (!handle) { 
      alert('Kan inte skapa ett XMLHTTP objekt');
      return false;
   }

   return handle;
} 

