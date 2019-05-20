
// JavaScript
window.sr = ScrollReveal();

// sr.reveal('h1', {
//     delay: 0,
//     duration: 200,
//     origin: 'bottom',
//     distance: '100px'
// });

function showNav(){
  console.log("ayyy lmao");
  var x = document.getElementById("responsive-nav");
  if(x.className === "responsive-nav"){
    x.className += " unfold";

  }
  else{
    x.className = "responsive-nav"
  }

}
