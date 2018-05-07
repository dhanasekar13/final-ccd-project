var ab=require('../js/try1')
var val=require('../js/validateexcel')

var value=''
val.validate().then(function(result){ value= result})


console.log(value)
