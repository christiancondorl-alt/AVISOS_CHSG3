// Identificadores públicos de configuración. NUNCA coloque una contraseña o clientSecret.
export const CONFIG = {
  mode: 'demo', // Cambiar a 'sharepoint' después de completar la instalación.
  tenantId: '',
  clientId: '',
  siteId: '',
  avisosListId: '',
  configuracionListId: '',
  configuracionItemId: '1',
  evidenciasDriveId: '',
  // Debe coincidir EXACTAMENTE con la URI SPA registrada en Microsoft Entra ID.
  redirectUri: '', // Ejemplo: https://USUARIO.github.io/avisos-sg3/auth.html
  scopes: ['Sites.ReadWrite.All']
};
