/* Shared verdict → tier mapping.
 * Single source of truth for turning a verdict/tier string ("Strong overall
 * fit", "Good Match", "Partial", "Low Match", …) into the CSS tier class that
 * carries its colours (see the .verdict-* rules in styles.css). Used by the
 * match-result page, the profile page, and the JD top-matches page so the
 * Strong/Good/Partial/Low colours are defined exactly once.
 */
(function (global) {
    function verdictTier(verdict) {
        var v = String(verdict == null ? '' : verdict).toLowerCase();
        if (v.indexOf('strong') !== -1) return 'strong';
        if (v.indexOf('good') !== -1) return 'good';
        if (v.indexOf('partial') !== -1) return 'partial';
        return 'low';
    }
    // Returns the pair of classes to put on a container: the base `.verdict`
    // (defines the custom-prop fallbacks) plus the tier-specific override.
    function verdictClass(verdict) {
        return 'verdict verdict-' + verdictTier(verdict);
    }
    global.verdictTier = verdictTier;
    global.verdictClass = verdictClass;
})(window);
