// Cierre de sesión por inactividad (10 minutos). No aplica a administradores.
(function() {
  var TIMEOUT_MS = 10 * 60 * 1000;
  var AVISO_MS = 30 * 1000;

  function esAdmin() {
    try {
      var emp = JSON.parse(localStorage.getItem('empleado'));
      return emp && (emp.admin === true || emp.admin === 'TRUE' || emp.admin === 'true');
    } catch(e) {
      return false;
    }
  }
  if (esAdmin()) return;

  var timer = null;
  var countdownInterval = null;
  var overlay = null;

  function cerrarSesionAuto() {
    localStorage.removeItem('empleado');
    window.location.href = 'index.html';
  }

  function quitarAviso() {
    if (countdownInterval) { clearInterval(countdownInterval); countdownInterval = null; }
    if (overlay) { overlay.parentNode.removeChild(overlay); overlay = null; }
  }

  function mostrarAviso() {
    if (overlay) return;
    var restante = AVISO_MS;
    overlay = document.createElement('div');
    overlay.className = 'sesion-overlay';
    overlay.innerHTML =
      '<div class="sesion-aviso">' +
        '<div class="sesion-aviso-titulo">Sesión por expirar</div>' +
        '<div class="sesion-aviso-texto">Por inactividad, la sesión se cerrará en <b id="sesionCuenta">30</b> segundos.</div>' +
        '<button id="sesionQuedarse">Seguir conectado</button>' +
      '</div>';
    document.body.appendChild(overlay);

    var numEl = document.getElementById('sesionCuenta');
    document.getElementById('sesionQuedarse').addEventListener('click', function() { reiniciar(); });

    countdownInterval = setInterval(function() {
      restante -= 1000;
      if (restante <= 0) { cerrarSesionAuto(); return; }
      if (numEl) numEl.textContent = Math.ceil(restante / 1000);
    }, 1000);

    timer = setTimeout(cerrarSesionAuto, AVISO_MS);
  }

  function reiniciar() {
    quitarAviso();
    clearTimeout(timer);
    timer = setTimeout(mostrarAviso, TIMEOUT_MS - AVISO_MS);
  }

  ['mousemove', 'mousedown', 'keydown', 'touchstart', 'scroll'].forEach(function(ev) {
    document.addEventListener(ev, reiniciar, { passive: true });
  });

  reiniciar();
})();
