
GUID=45d93b79-52d5-4ee8-bfba-ee4816bf0080

info:
	@ECHO "Utilisez 'make remove' pour supprimer le plugin de la base de registre de Rhino."

remove:
	dotnet script delete_registry.csx
