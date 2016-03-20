$(function() {
    $('a.page-scroll').bind('click', function(event) {
        $('html,body').animate({
            scrollTop: $(this.hash).offset().top
        }, 600);
        return false;
    });

    var h = $(window).height();

    $('header').css('height', h + 'px');

    new WOW().init();

    $(".owl-carousel").owlCarousel({
        items: 1,
        loop: true,
        autoplay: true,
        autoplaySpeed: 100,
        autoplayTimeout: 5000,
        nav: true,
        navText: ['<i class="fa fa-angle-left"></i>', '<i class="fa fa-angle-right"></i>'],
        navSpeed: 3000
    });
});

$('body').scrollspy({
    target: '.navbar-fixed-top'
});

$('.navbar-collapse ul li a').click(function() {
    $('.navbar-toggle:visible').click();
});

$(window).load(function() {
    var $container = $('.portfolioContainer');
    $('.grid').isotope({
        itemSelector: '.single-portfolio',
        layoutMode: 'fitRows'
    });
    $('.portfolio-item a').click(function() {
        $('.portfolio-item .current').removeClass('current');
        $(this).addClass('current');
        var selector = $(this).attr('data-filter');
        $container.isotope({
            filter: selector,
            animationOptions: {
                duration: 750,
                easing: 'linear',
                queue: false
            }
        });
        return false;
    });
});
