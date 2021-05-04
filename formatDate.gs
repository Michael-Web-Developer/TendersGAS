function formatDate(date, full = undefined, time = undefined) {
  if(date instanceof Date != true) return;
  var dd = date.getDate();
  if (dd < 10) dd = '0' + dd;

  var mm = date.getMonth() + 1;
  if (mm < 10) mm = '0' + mm;

  var yy = date.getFullYear();
  if (yy < 10) yy = '0' + yy;
  
  let returnDate = '';
  if(full == null){
    returnDate += mm + '.' + yy;
  } else {
    returnDate += dd + '.' + mm + '.' + yy;
  }
  
  if(time != null){
    var hh = date.getHours();
    if (hh < 10) hh = '0' + hh;
    
    var ii = date.getMinutes();
    if (ii < 10) ii = '0' + ii;
    
    returnDate += ' ' + hh + ':' + ii;   
  }
  
  return returnDate;
}