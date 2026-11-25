// Koppeling met Microsoft 365. Vul hier de waarden in uit de app-registratie in Microsoft Entra.
// Zolang clientId leeg is, draait de app in demomodus (lokale testgegevens, geen OneDrive, geen mail).
window.VDP_CONFIG = {
  clientId: '6308eecd-7e65-4f6d-beb9-579b4b99c16d',   // "Toepassings-id (client-id)" uit Entra
  tenantId: 'a94663cb-4f71-4335-8199-d25d6b8096cf',   // "Map-id (tenant-id)" uit Entra
  dataFolder: 'Apps/VdP Tijdregistratie',   // map in OneDrive waar uren.json bewaard wordt
  defaultRoots: ['Projecten 2026'],         // projectmappen in de hoofdmap van OneDrive
  rootPattern: '^Projecten 20\\d\\d$'       // nieuwe mappen met deze naam worden voorgesteld
};
