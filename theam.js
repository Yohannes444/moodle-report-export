var do_global_customizations = function() {
    if (ENV.current_user_roles.indexOf('student') > -1) {
        var globalNav = document.querySelector('.ic-app-nav-toggle-and-crumbs');
        if (globalNav) {
            globalNav.style.display = 'none';
        }
        var courseNav = document.querySelector('.ic-app-course-menu');
        if (courseNav) {
            courseNav.style.display = 'none';
        }
        var mainContent = document.getElementById('main');
        if (mainContent) {
            mainContent.style.marginLeft = '0';
            mainContent.style.width = '100%';
        }
    }
};
if (document.readyState === "complete" || (document.readyState !== "loading" && !document.documentElement.doScroll)) {
    do_global_customizations();
} else {
    document.addEventListener("DOMContentLoaded", do_global_customizations);
}