// Client-side password strength meter — mirrors _validate_password_complexity()'s
// exact 5 rules server-side (length >= PASSWORD_MIN_LENGTH, lowercase, uppercase,
// digit, special char) so this can never show "Strong" for something the server
// would actually reject. MIN_LENGTH below matches PASSWORD_MIN_LENGTH's default
// (8) — if that env var is overridden to something else, this meter's length
// check will be slightly out of sync with the server's, but the server-side
// validation (the actual gate) is unaffected either way.
(function () {
    var MIN_LENGTH = 8;

    function score(password) {
        var checks = [
            password.length >= MIN_LENGTH,
            /[a-z]/.test(password),
            /[A-Z]/.test(password),
            /\d/.test(password),
            /[^A-Za-z0-9]/.test(password),
        ];
        return checks.filter(Boolean).length;
    }

    var LEVELS = [
        {max: 2, label: 'Weak', color: '#dc2626'},
        {max: 4, label: 'Fair', color: '#d97706'},
        {max: 5, label: 'Strong', color: '#059669'},
    ];

    function levelFor(s) {
        for (var i = 0; i < LEVELS.length; i++) {
            if (s <= LEVELS[i].max) return LEVELS[i];
        }
        return LEVELS[LEVELS.length - 1];
    }

    window.attachPasswordStrength = function (inputId, meterId) {
        var input = document.getElementById(inputId);
        var meter = document.getElementById(meterId);
        if (!input || !meter) return;

        meter.innerHTML =
            '<div style="height:4px;border-radius:2px;background:var(--border);overflow:hidden;margin-top:6px;">' +
                '<div class="pw-strength-fill" style="height:100%;width:0%;transition:width .15s ease,background-color .15s ease;"></div>' +
            '</div>' +
            '<p class="pw-strength-label muted" style="font-size:0.78rem;margin-top:4px;min-height:1.1em;"></p>';
        var fill = meter.querySelector('.pw-strength-fill');
        var label = meter.querySelector('.pw-strength-label');

        function update() {
            var pw = input.value;
            if (!pw) {
                fill.style.width = '0%';
                label.textContent = '';
                return;
            }
            var s = score(pw);
            var level = levelFor(s);
            fill.style.width = (s / 5 * 100) + '%';
            fill.style.backgroundColor = level.color;
            label.textContent = 'Password strength: ' + level.label;
            label.style.color = level.color;
        }

        input.addEventListener('input', update);
        update();
    };
})();
