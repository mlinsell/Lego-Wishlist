var Config = (function () {

  const _REBRICKABLE_API_KEY = 'a462e39b3a0600b602afa264f1c07f2a';
  const _BRICKSET_API_KEY    = '3-iuiZ-56qL-yTuAU';

  function _require(name, value) {
    if (!value) throw new Error("Missing config value: " + name);
    return value;
  }

  return {
    rebrickableKey: () => _require('REBRICKABLE_API_KEY', _REBRICKABLE_API_KEY),
    bricksetKey:    () => _require('BRICKSET_API_KEY', _BRICKSET_API_KEY)
  };

})();
