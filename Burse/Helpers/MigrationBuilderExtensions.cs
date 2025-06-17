using Microsoft.EntityFrameworkCore.Migrations;
using System.Reflection;

namespace Burse.Helpers
{
    public static class MigrationBuilderExtensions
    {
        /// <summary>
        /// Execută un script SQL din embedded resource
        /// </summary>
        /// <param name="migrationBuilder">Migration builder instance</param>
        /// <param name="sqlFileName">Numele fișierului SQL (ex: "SeedData.sql")</param>
        public static void SqlResource(this MigrationBuilder migrationBuilder, string sqlFileName)
        {
            var assembly = Assembly.GetExecutingAssembly();

            // Caută resource-ul în assembly
            string resourceName = null;
            var resourceNames = assembly.GetManifestResourceNames();

            foreach (var name in resourceNames)
            {
                if (name.EndsWith(sqlFileName))
                {
                    resourceName = name;
                    break;
                }
            }

            if (resourceName == null)
            {
                throw new FileNotFoundException($"Embedded resource '{sqlFileName}' not found. Available resources: {string.Join(", ", resourceNames)}");
            }

            using (var stream = assembly.GetManifestResourceStream(resourceName))
            using (var reader = new StreamReader(stream))
            {
                var sqlScript = reader.ReadToEnd();
                migrationBuilder.Sql(sqlScript);
            }
        }

        /// <summary>
        /// Alternativă - citește direct din fișier (dacă nu vrei embedded resource)
        /// </summary>
        /// <param name="migrationBuilder">Migration builder instance</param>
        /// <param name="sqlFilePath">Calea către fișierul SQL</param>
        public static void SqlFile(this MigrationBuilder migrationBuilder, string sqlFilePath)
        {
            if (!File.Exists(sqlFilePath))
            {
                throw new FileNotFoundException($"SQL file not found: {sqlFilePath}");
            }

            var sqlScript = File.ReadAllText(sqlFilePath);
            migrationBuilder.Sql(sqlScript);
        }
    }

}
