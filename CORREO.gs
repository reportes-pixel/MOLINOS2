function forzarPermisosCorreo() {
  // Esta función no la vas a usar en tus botones, solo sirve para activar el permiso
  MailApp.sendEmail("test@example.com", "Permiso", "Validando scopes");
}