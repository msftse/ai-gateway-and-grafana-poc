output "id" {
  description = "Resource ID of the Content Safety account."
  value       = azurerm_cognitive_account.this.id
}

output "name" {
  description = "Name of the Content Safety account."
  value       = azurerm_cognitive_account.this.name
}

output "endpoint" {
  description = "Endpoint of the Content Safety account (used as the APIM content-safety backend URL)."
  value       = azurerm_cognitive_account.this.endpoint
}

output "primary_access_key" {
  description = "Primary access key of the Content Safety account (stored as an APIM named value)."
  value       = azurerm_cognitive_account.this.primary_access_key
  sensitive   = true
}
