document.addEventListener("DOMContentLoaded", () => {
  const menuBtn = document.getElementById("menuToggle");
  const mobileMenu = document.getElementById("mobileMenu");

  if (menuBtn && mobileMenu) {
    menuBtn.addEventListener("click", () => {
      mobileMenu.classList.toggle("show");
    });
  }

    // College Slider
    const collegeSlides = document.querySelectorAll(".college-slide");
    let collegeIndex = 0;

    if (collegeSlides.length > 0) {
        setInterval(() => {
            collegeSlides[collegeIndex].classList.remove("active");

            collegeIndex = (collegeIndex + 1) % collegeSlides.length;

            collegeSlides[collegeIndex].classList.add("active");
        }, 3500);
    }

});