using CsvHelper;
using CutQueue.Lib.Fabplan;
using CutQueue.Lib.Import.Model;
using CutQueue.Logging;
using System;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace CutQueue.Lib.JobImporter
{
    /// <summary>
    /// A utility that handles the transfer of orders from semi-colon 
    /// separated value files on the SIA server to the Fabplan server. 
    /// </summary>
    public static class JobImporter
    {
        private const string CSV_DATE_FORMAT = "yyyy'-'MM'-'dd";
        private static readonly NumberFormatInfo converterNumberFormatInfo = new NumberFormatInfo
        {
            NumberDecimalSeparator = ConfigINI.GetInstance().Items["NUMBER_DECIMAL_SEPARATOR"].ToString()
        };

        /// <summary>
        /// Imports a job from a csv file and exports it to Fabplan if it is valid
        /// </summary>
        /// <param name="csvFileInfo">A FileInfo object that point the csv file to import</param>
        public static async Task ImportJobFromCSVFileAndExportToFabplan(FileInfo csvFileInfo)
        {
            Job job = ImportJobFromCSVFile(csvFileInfo);
            SanitizeJob(ref job);

            if(job != null)
            {
                await ExportJobToFabplan(job);
            }
        }

        /// <summary>
        /// Imports a job from a csv file
        /// </summary>
        /// <param name="fileInfo">
        /// A FileInfo object that points to the csv file to import
        /// </param>
        /// <exception cref="InvalidJobIdentifierException">
        /// Thrown when the job identifier is equivalent to an empty string.
        /// </exception>
        /// <exception cref="InvalidSectionIdentifierException">
        /// Thrown when a section identifier doesn't evaluate to an integer.
        /// </exception>
        private static Job ImportJobFromCSVFile(FileInfo fileInfo)
        {
            Job job = new Job();
            using (StreamReader streamReader = new StreamReader(fileInfo.FullName))
            using (CsvReader csvReader = new CsvReader(streamReader))
            {
                bool jobInitialized = false;
                csvReader.Configuration.HasHeaderRecord = true;
                csvReader.Configuration.Delimiter = ";";
                csvReader.Read();
                csvReader.ReadHeader();
                while (csvReader.Read())
                {
                    if (!jobInitialized)
                    {
                        job = ExtractJobFromCsvEntry(in csvReader);
                        job.Customer = ExtractCustomerFromCsvEntry(csvReader);
                        job.Material = ExtractMaterialFromCsvEntry(csvReader);
                        if (string.IsNullOrWhiteSpace(job.Identifier))
                        {
                            string fileName = new Regex(@"\A(.*)(?:\.([^.]*)?)?\z").Match(fileInfo.Name.Trim()).Groups[1].Value;
                            if (!string.IsNullOrWhiteSpace(fileName))
                            {
                                job.Identifier = fileName.Trim();
                            }
                            else
                            {
                                throw new InvalidJobIdentifierException(fileName.Trim());
                            }
                        }
                        jobInitialized = true;
                    }

                    if (!int.TryParse(csvReader.GetField("Section's Identifier").Trim(), out int sectionIdentifier))
                    {
                        throw new InvalidSectionIdentifierException(csvReader.GetField("Section's Identifier"));
                    }

                    if (!job.JobTypeList.ContainsKey(sectionIdentifier))
                    {
                        job.JobTypeList.Add(sectionIdentifier, ExtractJobTypeFromCsvEntry(csvReader));
                    }

                    job.JobTypeList[sectionIdentifier].Parts.Add(ExtractPartFromCsvEntry(csvReader));
                }
            }

            return job;
        }

        /// <summary>
        /// Parses a Job object out of a csv entry
        /// </summary>
        /// <param name="csvReader">
        /// A CsvReader object that points to a csv entry
        /// </param>
        private static Job ExtractJobFromCsvEntry(in CsvReader csvReader)
        {
            Job job = new Job();

            if (!string.IsNullOrWhiteSpace(csvReader.GetField("Order's Identifier")))
            {
                job.Identifier = csvReader.GetField("Order's Identifier").Trim();
            }

            if (!string.IsNullOrWhiteSpace(csvReader.GetField("Order's Purchase Order Identifier")))
            {
                job.PONumber = csvReader.GetField("Order's Purchase Order Identifier").Trim();
            }
            else
            {
                job.PONumber = null;
            }

            string rawDate = csvReader.GetField("Order's Required Date").Trim();
            if (DateTime.TryParseExact(rawDate, CSV_DATE_FORMAT, null, DateTimeStyles.None, out DateTime requiredDate))
            {
                job.RequiredDate = requiredDate;
            }
            else
            {
                job.RequiredDate = null;
            }

            return job;
        }

        /// <summary>
        /// Parses a Customer object out of a csv entry
        /// </summary>
        /// <param name="csvReader">
        /// A CsvReader object that points to a csv entry
        /// </param>
        private static Customer ExtractCustomerFromCsvEntry(in CsvReader csvReader)
        {
            return new Customer()
            {
                Name = csvReader.GetField("Customer's Name").Trim(),
                Address1 = csvReader.GetField("Customer's Address Line 1").Trim(),
                Address2 = csvReader.GetField("Customer's Address Line 2").Trim(),
                PostalCode = csvReader.GetField("Customer's Postal Code").Trim()
            };
        }

        /// <summary>
        /// Parses a Material object out of a csv entry
        /// </summary>
        /// <param name="csvReader">
        /// A CsvReader object that points to a csv entry
        /// </param>
        private static Material ExtractMaterialFromCsvEntry(in CsvReader csvReader)
        {
            return new Material()
            {
                Essence = csvReader.GetField("Material's Essence").Trim(),
                Grade = csvReader.GetField("Material's Grade").Trim()
            };
        }

        /// <summary>
        /// Parses a JobType object out of a csv entry
        /// </summary>
        /// <param name="csvReader">
        /// A CsvReader object that points to a csv entry
        /// </param>
        private static JobType ExtractJobTypeFromCsvEntry(in CsvReader csvReader)
        {
            JobType jobType = new JobType();

            if (!string.IsNullOrWhiteSpace(csvReader.GetField("Section's Model")))
            {
                jobType.Model = csvReader.GetField("Section's Model").Trim();
            }
            else
            {
                jobType.Model = null;
            }

            if (int.TryParse(csvReader.GetField("Section's Type").Trim(), out int type))
            {
                jobType.Type = type;
            }
            else
            {
                jobType.Type = null;
            }

            if (!string.IsNullOrWhiteSpace(csvReader.GetField("Section's External Profile")))
            {
                jobType.ExternalProfile = csvReader.GetField("Section's External Profile").Trim();
            }
            else
            {
                jobType.ExternalProfile = null;
            }

            return jobType;
        }

        /// <summary>
        /// Parses a Part object out of a csv entry
        /// </summary>
        /// <param name="csvReader">
        /// A CsvReader object that points to a csv entry
        /// </param>
        /// <exception cref="InvalidPartQuantityException">
        /// Thrown when a part has a quantity that doesn't evaluate to an integer.
        /// </exception>
        /// <exception cref="InvalidPartGrainDirectionException">
        /// Thrown when a part has a grain direction that doesn't evaluate to an integer.
        /// </exception>
        /// <exception cref="InvalidPartHeightException">
        /// Thrown when a part has a height that doesn't evaluate to an integer.
        /// </exception>
        /// <exception cref="InvalidPartWidthException">
        /// Thrown when a part has a width that doesn't evaluate to an integer.
        /// </exception>
        private static Part ExtractPartFromCsvEntry(in CsvReader csvReader)
        {
            Part part = new Part();

            if (int.TryParse(csvReader.GetField("Part's Quantity").Trim(), out int partQuantity))
            {
                part.Quantity = partQuantity;
            }
            else
            {
                throw new InvalidPartQuantityException(csvReader.GetField("Part's Quantity"));
            }

            string grainDirection = csvReader.GetField("Part's Grain Direction").Trim();
            if (Array.Exists(new[] { "N", "X", "Y" }, (grain) => grainDirection == grain))
            {
                part.GrainDirection = grainDirection;
            }
            else
            {
                throw new InvalidPartGrainDirectionException(csvReader.GetField("Part's Grain Direction"));
            }

            string rawHeight = csvReader.GetField("Part's Height").Trim();
            if (decimal.TryParse(rawHeight, NumberStyles.AllowDecimalPoint, converterNumberFormatInfo, out decimal partHeight))
            {
                part.Height = partHeight;
            }
            else
            {
                throw new InvalidPartHeightException(csvReader.GetField("Part's Height"));
            }

            string rawWidth = csvReader.GetField("Part's Width").Trim();
            if (decimal.TryParse(rawWidth, NumberStyles.AllowDecimalPoint, converterNumberFormatInfo, out decimal partWidth))
            {
                part.Width = partWidth;
            }
            else
            {
                throw new InvalidPartWidthException(csvReader.GetField("Part's Width"));
            }

            return part;
        }

        /// <summary>
        /// Sanitizes an imported job by removing all parts with an invalid quantity or dimensions, removing empty jobTypes 
        /// and ultimately setting a job without jobTypes to null, disallowing its exportation.
        /// </summary>
        /// <param name="job">The job to sanitize</param>
        private static void SanitizeJob(ref Job job)
        {
            foreach (int i in job.JobTypeList.Keys)
            {
                JobType jobType = job.JobTypeList[i];
                for (int j = jobType.Parts.Count - 1; j >= 0; j--)
                {
                    Part part = jobType.Parts[j];
                    if (part.Quantity <= 0 || part.Height <= 0 || part.Width <= 0)
                    {
                        jobType.Parts.RemoveAt(j);
                    }
                }

                if (!jobType.Parts.Any())
                {
                    job.JobTypeList.RemoveAt(i);
                }
            }

            if (!job.JobTypeList.Any())
            {
                job = null;
            }
        }

        /// <summary>
        /// Exports the job to Fabplan.
        /// </summary>
        /// <param name="job">The job to send to Fabplan</param>
        /// <exception cref="CannotCreateJobException">Thrown when the job creation fails</exception>
        private static async Task ExportJobToFabplan(Job job)
        {
            /* Determine if the job exists. */
            bool exists = await JobExists(job);

            /* Determine if job is linked to a batch. If job doesn't exist, it is considered as not linked to any batch. */
            bool isLinked = false;
            if (exists)
            {
                isLinked = await JobIsLinked(job);
            }

            if (!isLinked)
            {
                /* Delete the existing job if it must still be imported. */
                if (exists)
                {
                    await DeleteJob(job);
                }

                try
                {
                    await CreateJob(job);
                }
                catch (CannotCreateJobException e)
                {
                    Logger.Log($"Failed to create job \"{job.Identifier}\": {e.ToString()}.");
                }
            }
        }

        /// <summary>
        /// Creates a job in Fabplan.
        /// </summary>
        /// <param name="job">The job to create on the server</param>
        /// <exception cref="Exception">Thrown when the creation fails</exception>
        private static async Task CreateJob(Job job)
        {
            UriBuilder builder = new UriBuilder()
            {
                Host = ConfigINI.GetInstance().Items["HOST_NAME"].ToString(),
                Path = ConfigINI.GetInstance().Items["CREATE_JOB_URL"].ToString(),
                Port = -1,
                Scheme = "http"
            };

            try
            {
                await FabplanHttpRequest.Post(builder.ToString(), job);
            }
            catch (FabplanHttpResponseWarningException e)
            {
                throw new CannotOverwriteJobException($"Job \"{job.Identifier}\" could not be overwritten.", e);
            }
            catch (Exception e)
            {
                throw new CannotCreateJobException($"Job \"{job.Identifier}\" could not be created.", e);
            }
        }


        /// <summary>
        /// Deletes a job from Fabplan
        /// </summary>
        /// <param name="job">The job to delete from the server (found by job identifier)</param>
        /// <exception cref="Exception">Thrown when the deletion fails</exception>
        private async static Task DeleteJob(Job job)
        {
            UriBuilder builder = new UriBuilder()
            {
                Host = ConfigINI.GetInstance().Items["HOST_NAME"].ToString(),
                Path = ConfigINI.GetInstance().Items["DELETE_JOB_URL"].ToString(),
                Port = -1,
                Scheme = "http"
            };

            try
            {
                await FabplanHttpRequest.Post(builder.ToString(), new { name = job.Identifier });
            }
            catch (Exception e)
            {
                throw new Exception($"Job \"{job.Identifier}\" could not be deleted.", e);
            }
        }


        /// <summary>
        /// Queries the server to determine if a job exists (found by job identifier).
        /// </summary>
        /// <param name="job">The Job Object to test for existence</param>
        /// <exception cref="Exception">Thrown when the request to the server fails</exception>
        /// <returns>True if job exists, false otherwise</returns>
        private async static Task<bool> JobExists(Job job)
        {
            UriBuilder builder = new UriBuilder()
            {
                Host = ConfigINI.GetInstance().Items["HOST_NAME"].ToString(),
                Path = ConfigINI.GetInstance().Items["JOB_EXISTS_URL"].ToString(),
                Port = -1,
                Scheme = "http"
            };

            bool exists = true;
            try
            {
                exists = await FabplanHttpRequest.Get(builder.ToString(), new { name = job.Identifier });
            }
            catch (Exception e)
            {
                throw new Exception($"Failed to determine if job \"{job.Identifier}\" exists.", e);
            }

            return exists;
        }

        /// <summary>
        /// Queries the server to determine if a job (found by job identifier) is linked to a batch.
        /// </summary>
        /// <param name="job">The job to test a link to a batch for</param>
        /// <exception cref="Exception">Thrown when the request to the server fails</exception>
        /// <returns>True if job is linked to a batch, false otherwise</returns>
        private async static Task<bool> JobIsLinked(Job job)
        {
            UriBuilder builder = new UriBuilder()
            {
                Host = ConfigINI.GetInstance().Items["HOST_NAME"].ToString(),
                Path = ConfigINI.GetInstance().Items["JOB_IS_LINKED_URL"].ToString(),
                Port = -1,
                Scheme = "http"
            };

            bool isLinked = true;
            try
            {
                isLinked = await FabplanHttpRequest.Get(builder.ToString(), new { name = job.Identifier });
            }
            catch (Exception e)
            {
                throw new Exception($"Failed to determine if job \"{job.Identifier}\" is linked to a batch.", e);
            }

            return isLinked;
        }
    }

    /// <summary>
    /// An exception thrown when a part's height is invalid. 
    /// </summary>
    public class InvalidPartHeightException : Exception
    {
        public InvalidPartHeightException(string partHeight) : base($"Invalid part height \"{partHeight}\".") { }
    }

    /// <summary>
    /// An exception thrown when a part's width is invalid. 
    /// </summary>
    public class InvalidPartWidthException : Exception
    {
        public InvalidPartWidthException(string partWidth) : base($"Invalid part width \"{partWidth}\".") { }
    }

    /// <summary>
    /// An exception thrown when a part's quantity is invalid. 
    /// </summary>
    public class InvalidPartQuantityException : Exception
    {
        public InvalidPartQuantityException(string partQuantity) : base($"Invalid part quantity \"{partQuantity}\".") { }
    }

    /// <summary>
    /// An exception thrown when a part's grain direction is invalid. 
    /// </summary>
    public class InvalidPartGrainDirectionException : Exception
    {
        public InvalidPartGrainDirectionException(string partGrainDirection) : 
            base($"Invalid part grain direction \"{partGrainDirection}\".") { }
    }

    /// <summary>
    /// An exception thrown when a job identifier is invalid. 
    /// </summary>
    public class InvalidJobIdentifierException : Exception
    {
        public InvalidJobIdentifierException(string jobIdentifier) : base($"Invalid job identifier \"{jobIdentifier}\".") { }
    }

    /// <summary>
    /// An exception thrown when a section identifier is invalid. 
    /// </summary>
    public class InvalidSectionIdentifierException : Exception
    {
        public InvalidSectionIdentifierException(string sectionIdentifier) : 
            base($"Invalid section identifier \"{sectionIdentifier}\".") { }
    }

    /// <summary>
    /// An exception thrown when a job cannot be overwritten. 
    /// </summary>
    public sealed class CannotOverwriteJobException : Exception
    {
        public CannotOverwriteJobException(string message, Exception innerException) : base(message, innerException) { }
    }

    /// <summary>
    /// An exception thrown when a job cannot be created. 
    /// </summary>
    public sealed class CannotCreateJobException : Exception
    {
        public CannotCreateJobException(string message, Exception innerException) : base(message, innerException) { }
    }
}