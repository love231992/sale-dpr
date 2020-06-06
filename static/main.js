const navSlide = () => {
    const burger = document.querySelector('.burger');
    const nav = document.querySelector('.navbar');
    const navLinks = document.querySelectorAll('.navbar li')
    burger.addEventListener('click', () => {
        nav.classList.toggle('nav-active');
        navLinks.forEach((links,index) => {
            if (links.style.animation) {
                links.style.animation = "";
            } else {
            links.style.animation = `navLinkFade 0.5s ease forwards ${index / 7 + 0.5}s`;
        }
        });
        burger.classList.toggle('toggle');
    })
    
}

navSlide();



document.getElementById('btn-yester').addEventListener('click', () => {
    document.getElementById('btn-yester').style.backgroundColor = "#3d0c11";
    document.getElementById('btn-mon').style.backgroundColor = "#DF2935";
    document.getElementById('btn-tod').style.backgroundColor = "#DF2935";
    document.getElementById('hello').innerHTML = document.getElementById('yes').innerHTML
})
document.getElementById('btn-tod').addEventListener('click', () => {
    document.getElementById('btn-yester').style.backgroundColor = "#DF2935";
    document.getElementById('btn-mon').style.backgroundColor = "#DF2935";
    document.getElementById('btn-tod').style.backgroundColor = "#3d0c11";
    document.getElementById('hello').innerHTML = document.getElementById('dummy').innerHTML
})
document.getElementById('btn-mon').addEventListener('click', () => {
    document.getElementById('btn-yester').style.backgroundColor = "#DF2935";
    document.getElementById('btn-mon').style.backgroundColor = "#3d0c11";
    document.getElementById('btn-tod').style.backgroundColor = "#DF2935";
    document.getElementById('hello').innerHTML = document.getElementById('month').innerHTML
})


a = document.querySelector("body > div.sale > div.balance > table > tbody > tr:nth-child(2) > td:nth-child(3)")
console.log(a)
y = document.querySelector('.balance')
console.log(y)
x = document.querySelector("body > div.sale > div.balance > table > tbody > tr:nth-child(3) > td:nth-child(3)")
console.log(x)



