using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using Dapper;
using System.Data.SqlClient;
using Microsoft.SqlServer.Server;

namespace TableNaker
{
    public static class SQLBuilder
    {
        public static StringBuilder CheckForExisting(
            this StringBuilder sb,
            string tableName,
            ref DynamicParameters parameter)
        {
            sb.Append($@" IF OBJECT_ID(@TableName) IS NOT NULL
                          BEGIN
                              SELECT 1
                          END
                          ELSE BEGIN
                              SELECT 0
                          END
                    ");

            parameter.Add("@TableName", tableName);

            return sb;
        }

        public static StringBuilder RemoveExistingTable(
            this StringBuilder sb,
            string tableName,
            ref DynamicParameters parameter)
        {
            sb.Append($@" IF OBJECT_ID(@TableNameRemove) IS NOT NULL
                          BEGIN
                              DROP TABLE [dbo].[{tableName}]
                          END
                    ");

            parameter.Add("@TableNameRemove", tableName);

            return sb;
        }

        public static StringBuilder AppendColumnHeaders(
            this StringBuilder sb, 
            string tableName, 
            ref Dictionary<string,string> headersInOrder)
        {
            sb.Append($" CREATE TABLE  [dbo].[" + tableName + "] ( ");

            CreateHeaderInsertion(ref sb, ref headersInOrder);

            sb.Append(" ) ");

            return sb;
        }

        private static void CreateHeaderInsertion(ref StringBuilder sb, ref Dictionary<string, string> headersInOrder, bool excludeTypeDefinitions = false)
        {
            if (excludeTypeDefinitions)
            {
                sb.Append(

                    string.Join(
                        ",",
                        headersInOrder.Select(header => $" [{header.Key}] ")

                    )

                );
            }
            else
            {
                sb.Append(

                        string.Join(
                            ",",
                            headersInOrder.Select(header => $" [{header.Key}] {header.Value} ")

                        )

                    );
            }
        }

        public static StringBuilder AppendTVPRemove(
            this StringBuilder sb,
            string tableName)
        {
            sb.Append($" REMOVE TYPE [dbo].[tvp" + tableName + "] ");
            return sb;
        }

        public static StringBuilder AppendColumnHeadersTVP(
            this StringBuilder sb,
            string tableName,
            ref Dictionary<string, string> headersInOrder)
        {
            sb.Append($" CREATE TYPE [dbo].[tvp" + tableName + "] AS TABLE ( ");

            CreateHeaderInsertionTVP(ref sb, ref headersInOrder);

            sb.Append(" ) ");

            return sb;
        }

        private static void CreateHeaderInsertionTVP(ref StringBuilder sb, ref Dictionary<string, string> headersInOrder, bool excludeTypeDefinitions = false)
        {
            if (excludeTypeDefinitions)
            {
                sb.Append(

                    string.Join(
                        ",",
                        headersInOrder.Select(header => $" [{header.Key}] ")

                    )

                );
            }
            else
            {
                sb.Append(

                    string.Join(
                        ",",
                        headersInOrder.Select(header => $" [{header.Key}] {header.Value} ")

                    )

                );
            }
        }

        //https://gist.github.com/divega/f0f88bf16f35641239cfd9bc534e8d7c
        public static StringBuilder AddBatch(
            this StringBuilder sb,
            string tableName,
            ref List<string[]> rows,
            ref Dictionary<string, string> headersInOrder,
            ref DynamicParameters parameters)
        {
            sb.Append($" INSERT INTO [dbo].[{tableName}] ( ");

            CreateHeaderInsertion(ref sb, ref headersInOrder, true);

            sb.Append(" ) ");
            sb.Append(" VALUES ");

            var outerRowCount = rows.Count;
            var outerIndexer = 1;
            var gandIndexer = 0;
            foreach (var row in rows)
            {
                var rowCount = row.Length;
                var indexer = 1;
                sb.Append(" ( ");

                foreach (var subRow in row)
                {
                    sb.Append($"@{gandIndexer}");

                    parameters.Add($"@{gandIndexer}", subRow);

                    if (indexer < rowCount)
                        sb.Append(" , ");
                  
                    indexer++;
                    gandIndexer++;
                }

                sb.Append(" ) ");

                if (outerIndexer < outerRowCount)
                    sb.Append(" , ");

                outerIndexer++;
            }

            return sb;
        }
    }
}