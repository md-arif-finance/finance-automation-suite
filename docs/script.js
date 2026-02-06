document.addEventListener('DOMContentLoaded', () => {
    // Console log for now, can add interactivity later
    console.log('Automated AR Control System Portfolio Loaded');

    // Modal Interaction Logic
    const getTemplateBtn = document.getElementById('get-template-btn');
    const modalOverlay = document.getElementById('verification-modal');
    const modalConfirmBtn = document.getElementById('modal-confirm-btn');

    if (getTemplateBtn && modalOverlay && modalConfirmBtn) {
        getTemplateBtn.addEventListener('click', (e) => {
            e.preventDefault();
            modalOverlay.classList.remove('d-none');
            // Small delay to allow CSS transition to work
            setTimeout(() => {
                modalOverlay.classList.add('active');
            }, 10);
        });
    }

    // Mobile Menu Toggle
    const mobileMenuBtn = document.querySelector('.mobile-menu-btn');
    const nav = document.querySelector('nav');

    if (mobileMenuBtn && nav) {
        mobileMenuBtn.addEventListener('click', () => {
            nav.classList.toggle('active');
        });

        // Close menu when a link is clicked
        const navLinks = nav.querySelectorAll('a');
        navLinks.forEach(link => {
            link.addEventListener('click', () => {
                nav.classList.remove('active');
            });
        });
    }

    // Confirm Action
    if (modalConfirmBtn) {
        modalConfirmBtn.addEventListener('click', () => {
            // Redirect to the actual Google Sheet in a new tab
            window.open("https://docs.google.com/spreadsheets/d/1s0hWRnYPF35DMEu5q2WZuUQxisJscrmzIXuU1Nm7LyI/copy", "_blank");
            // Close the modal
            modalOverlay.classList.remove('active');
            setTimeout(() => {
                modalOverlay.classList.add('d-none');
            }, 300);
        });

        // Close on outside click (Optional, but good UX)
        modalOverlay.addEventListener('click', (e) => {
            if (e.target === modalOverlay) {
                modalOverlay.classList.remove('active');
                setTimeout(() => {
                    modalOverlay.classList.add('d-none');
                }, 300); // Match CSS transition duration
            }
        });
    }



    // Load YouTube Iframe API
    const tag = document.createElement('script');
    tag.src = "https://www.youtube.com/iframe_api";
    const firstScriptTag = document.getElementsByTagName('script')[0];
    firstScriptTag.parentNode.insertBefore(tag, firstScriptTag);

    // Video Container Interaction
    const videoContainer = document.getElementById('video-container');
    let player;
    let originalContent = '';

    if (videoContainer) {
        // Store the original content (image + overlay + button)
        originalContent = videoContainer.innerHTML;

        videoContainer.addEventListener('click', function () {
            // Prevent multiple players
            if (videoContainer.classList.contains('playing')) return;

            videoContainer.classList.add('playing');
            // Inject spinner and player placeholder
            videoContainer.innerHTML = `
                <div class="loading-spinner" style="display: block;"></div>
                <div id="yt-player"></div>
            `;

            player = new YT.Player('yt-player', {
                height: '100%',
                width: '100%',
                videoId: 'Ye4R67UrsFs',
                host: 'https://www.youtube-nocookie.com', // Privacy Mode
                playerVars: {
                    'autoplay': 1,
                    'rel': 0,
                    'modestbranding': 1
                },
                events: {
                    'onReady': onPlayerReady,
                    'onStateChange': onPlayerStateChange
                }
            });
        });
    }

    function onPlayerReady(event) {
        // Hide spinner when video is ready
        const spinner = videoContainer.querySelector('.loading-spinner');
        if (spinner) {
            spinner.style.display = 'none';
        }
    }

    function onPlayerStateChange(event) {
        // YT.PlayerState.ENDED is 0
        if (event.data === 0) {
            // Video ended
            resetVideoContainer();
        }
    }

    function resetVideoContainer() {
        if (player) {
            player.destroy();
            player = null;
        }
        if (videoContainer) {
            videoContainer.innerHTML = originalContent;
            videoContainer.classList.remove('playing');
        }
    }
});
