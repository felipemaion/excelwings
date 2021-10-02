var btnRent = function(){var a=["https://babycentral.com.br/aluguel/domestico/aluguel-de-bombinha-tira-leite",
"https://babycentral.com.br/aluguel/domestico/alugar-tira-leite",
"https://babycentral.com.br/aluguel/domestico/alugar-bombinha-tira-leite-pump-in-style",
"https://babycentral.com.br/aluguel/domestico/aluguel-de-bombinha-de-tirar-leite",
"https://babycentral.com.br/aluguel/domestico/aluguel-de-bombinha-tira-leite-medela-swing-maxi",
"https://babycentral.com.br/extrator-de-leite/aluguel/aluguel-extrator-de-leite-lactina-medela",
"https://www.babycentral.com.br/extrator-de-leite/aluguel/aluguel-extrator-de-leite-lactina-medela",
"https://www.babycentral.com.br/aluguel/domestico/aluguel-de-bombinha-tira-leite-medela-swing-maxi",
"https://www.babycentral.com.br/aluguel/domestico/aluguel-de-bombinha-tira-leite",
"https://www.babycentral.com.br/aluguel/domestico/alugar-tira-leite",
"https://www.babycentral.com.br/aluguel/domestico/alugar-bombinha-tira-leite-pump-in-style",
"https://www.babycentral.com.br/aluguel/domestico/aluguel-de-bombinha-de-tirar-leite"],
e = !1,
t=document.location.href;
for(var i in a)
    if(-1 !== t.indexOf(a[i])){e=!0;break}
    e && (setInterval(function(){
            $(".botao-comprar.principal.grande").html(
                '<i class="icon-shopping-cart"><svg class="icon" xmlns="http://www.w3.org/2000/svg" viewBox="-8173.381 616 12.643 18.48"><path class="a" d="M15.062,7.16h-.79V5.4A4.2,4.2,0,0,0,10.321,1,4.2,4.2,0,0,0,6.37,5.4V7.16H5.58A1.682,1.682,0,0,0,4,8.92v8.8a1.682,1.682,0,0,0,1.58,1.76h9.482a1.682,1.682,0,0,0,1.58-1.76V8.92A1.682,1.682,0,0,0,15.062,7.16Zm-4.741,7.92a1.682,1.682,0,0,1-1.58-1.76,1.682,1.682,0,0,1,1.58-1.76,1.682,1.682,0,0,1,1.58,1.76A1.682,1.682,0,0,1,10.321,15.08Zm2.449-7.92h-4.9V5.4a2.6,2.6,0,0,1,2.449-2.728A2.6,2.6,0,0,1,12.771,5.4Z" transform="translate(-8177.381 615)"></path></svg></i>Quero Alugar'),
                $('.botao-comprar.principal.grande')
                    .attr('href', '#noexists'),
                    $('.accordion .accordion-group').eq(1).remove()
                },10),
                $(".botao-comprar.principal.grande")
                    .attr("href","#noexists"),
                    $(".botao-comprar.principal.grande")
                        .click(function(){
                            -1 !== t.indexOf("babycentral.com.br/aluguel/domestico/alugar-bombinha-tira-leite-pump-in-style")
                            ? window.location.href = "https://www.babycentral.com.br/pagina/bombinha-de-tirar-leite-medela-pump-in-style.html" 
                            : -1!==t.indexOf("babycentral.com.br/aluguel/domestico/aluguel-de-bombinha-de-tirar-leite")
                            ? window.location.href = "https://www.babycentral.com.br/pagina/bombinha-tira-leite-medela-swing.html"
                            : -1!==t.indexOf("babycentral.com.br/aluguel/domestico/aluguel-de-bombinha-tira-leite-medela-swing-maxi")
                            ? window.location.href = "https://www.babycentral.com.br/pagina/bombinha-tira-leite-eletrica-medela-swing-maxi.html" 
                            : -1 !== t.indexOf("babycentral.com.br/extrator-de-leite/aluguel/aluguel-extrator-de-leite-lactina-medela") && window.location.href = "https://www.babycentral.com.br/pagina/extrator-de-leite-materno-eletrico-lactina-medela.html"
                        }))};

try{
	$(document).ready(function () {
		btnRent();
	});
} catch(err) {
	console.log("document ready broken ", err);
	btnRent();
}