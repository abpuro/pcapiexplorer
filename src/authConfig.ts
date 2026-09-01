export const partnerCenterAuthConfig = {
  clientId: import.meta.env.VITE_PARTNER_CENTER_CLIENT_ID ?? "",
  tenantId: import.meta.env.VITE_PARTNER_CENTER_TENANT_ID ?? "organizations",
  scopes: [
    import.meta.env.VITE_PARTNER_CENTER_SCOPE ??
      "https://api.partnercenter.microsoft.com/user_impersonation",
  ],
  baseUrl: import.meta.env.VITE_PARTNER_CENTER_BASE_URL ?? "https://api.partnercenter.microsoft.com/v1",
};
