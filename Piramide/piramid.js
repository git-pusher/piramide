var nivelNode = document.getElementById('nivel');
var rawCaracter = document.getElementById('caracter');
// var rawColor = document.getElementById('color');
var result = document.getElementById('piramide');
var nivel, caracter;
// var linea = ' ';

//forma uno: declararlo como una función
function getNivel () {
    nivel = nivelNode.value
}
//forma dos: guardarlo en una variable
getCaracter = () => {
    caracter = rawCaracter.value
}
//forma tres: como un arrow function
// getColor = () => {
//     color = rawColor.value
// }

nivelNode.addEventListener('keyup', getNivel)

createPyramid = () => {
    
    linea = ' ';
    piramide.innerHTML = linea;
    getCaracter();
    // color = rawColor.value;
    for (let i = 0; i < nivel; i++){
        piramide.innerHTML += '<p class="">' + (linea += ('<span class="">' + caracter + '</span>')) + '</p>';
    }
}