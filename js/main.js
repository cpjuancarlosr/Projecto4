document.addEventListener('DOMContentLoaded', () => {
    const contactForm = document.getElementById('contactForm');
    if (contactForm) {
        contactForm.addEventListener('submit', (e) => {
            e.preventDefault();
            const formContainer = contactForm.parentNode;
            formContainer.innerHTML = '<h2>Thank you for your message!</h2><p>We will get back to you soon.</p>';
        });
    }
});
