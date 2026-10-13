(function () {
    var DESIGN_W = 1536;
    var DESIGN_H = 1024;

    function fit() {
        var canvas = document.querySelector('.login-canvas');
        if (!canvas) return;
        // "Cover" scaling (like background-size:cover), not "contain" — the
        // canvas always fills the whole viewport edge to edge with no
        // letterbox gap on either side, at the cost of cropping whichever
        // dimension overflows. .login-viewport's overflow:hidden clips it.
        var scale = Math.max(window.innerWidth / DESIGN_W, window.innerHeight / DESIGN_H);
        canvas.style.transform = 'scale(' + scale.toFixed(4) + ')';
    }

    window.addEventListener('resize', fit);
    window.addEventListener('load', fit);
    if (document.readyState === 'loading') {
        document.addEventListener('DOMContentLoaded', fit);
    } else {
        fit();
    }
})();
