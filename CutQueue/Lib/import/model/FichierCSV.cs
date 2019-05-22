using CutQueue.Lib.Fabplan;
using System;
using System.Collections.Generic;
using System.Dynamic;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace CutQueue.Lib.import.model
{
    /// <summary>
    /// A class that reads and imports a csv job file from SIA ProPlan to Fabplan
    /// </summary>
    class FichierCSV
    {

        private FileInfo _file;
	    private List<EntreeCSV> _entrees;
	
        /// <summary>
        /// Constructor that accepts a file as an argument
        /// </summary>
        /// <param name="file">The csv job file</param>
	    public FichierCSV(FileInfo file)
        {
		    _file = file;
            _entrees = new List<EntreeCSV>();
		
		    ReadFile();
        }

        /// <summary>
        /// Parses the csv file into a list of EntreeCSV objects
        /// </summary>
        private void ReadFile()
        {
            var fileStream = new FileStream(_file.FullName, FileMode.Open, FileAccess.Read);

            using (StreamReader streamReader = new StreamReader(fileStream, Encoding.UTF8))
            {
                string line;
                while ((line = streamReader.ReadLine()) != null)
                {
                    EntreeCSV entree = new EntreeCSV(line);
                    if (entree.Hauteur.Trim() != "")
                    {
                        _entrees.Add(entree);
                    }
                }
            }
        }


        /// <summary>
        /// Converts the list of EntreeCSV into a job object that is then sent to Fabplan for creation.
        /// </summary>
        public async Task Import()
        {
            /* Do not create a job that contains no part. */
            if (_entrees.Count > 0)
            {
                /* Determine if the job exists. */
                bool exists = await JobExists();

                /* Determine if job is linked to a batch. If job doesn't exist, it is considered as not linked to any batch. */
                bool isLinked = false;
                if(exists)
                {
                    isLinked = await JobIsLinked();
                }

                if (!isLinked)
                {
                    /* Delete the existing job if it must still be imported. */
                    if (exists)
                    {
                        await DeleteJob();
                    }

                    dynamic job = new ExpandoObject();
                    job.name = GetJobNumber();
                    job.deliveryDate = (_entrees[0].DueDate != "" && _entrees[0].DueDate != null) ? _entrees[0].DueDate : null;
                    job.jobTypes = new List<ExpandoObject>();

                    foreach (EntreeCSV entree in _entrees)
                    {
                        /* If the jobtype of the current part is not the same as the jobtype of the previous part, create a new jobtype. */
                        dynamic jobType = null;
                        int count = job.jobTypes.Count;
                        int index = count - 1;
                        string type = (index >= 0) ? job.jobTypes[index].type : null;
                        string model = (index >= 0) ? job.jobTypes[index].model : null;
                        string externalProfile = (index >= 0) ? job.jobTypes[index].externalProfile : null;
                        if (count <= 0 || entree.Type != type || entree.Modele != model || entree.Profil != externalProfile)
                        {
                            jobType = new ExpandoObject();
                            jobType.type = entree.Type;
                            jobType.model = entree.Modele;
                            jobType.externalProfile = entree.Profil;
                            jobType.material = entree.Essence;
                            jobType.parts = new List<ExpandoObject>();
                            job.jobTypes.Add(jobType);
                        }
                        jobType = job.jobTypes[job.jobTypes.Count - 1];


                        /* Do not create inputs with no dimension. */
                        if (entree.Hauteur != "" && entree.Largeur != "")
                        {
                            dynamic part = new ExpandoObject();
                            part.quantity = entree.Quantite;
                            part.height = entree.Hauteur;
                            part.width = entree.Largeur;
                            part.grain = entree.Grain;
                            jobType.parts.Add(part);
                        }
                    }

                    try
                    {
                        await CreateJob(job);
                    }
                    catch (CannotOverwriteJobException)
                    {
                        /* The job could not be updated. */
                    }
                }
            }
        }

        /// <summary>
        /// Creates a job in Fabplan.
        /// </summary>
        /// <param name="job">An object that represents the job to create</param>
        /// <exception cref="Exception">Thrown when the creation fails</exception>
        private async Task CreateJob(object job)
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
                throw new CannotOverwriteJobException($"Job \"{GetJobNumber()}\" could not be overwritten.", e);
            }
            catch (Exception e)
            {
                throw new CannotCreateJobException($"Job \"{GetJobNumber()}\" could not be created.", e);
            }
        }


        /// <summary>
        /// Deletes a job from Fabplan
        /// </summary>
        /// <exception cref="Exception">Thrown when the deletion fails</exception>
        private async Task DeleteJob()
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
                await FabplanHttpRequest.Post(builder.ToString(), new { name = GetJobNumber() });
            }
            catch (Exception e)
            {
                throw new Exception($"Job \"{GetJobNumber()}\" could not be deleted.", e);
            }
        }


        /// <summary>
        /// Returns the name of a job (its production number)
        /// </summary>
        /// <returns>The name of the job</returns>
        private string GetJobNumber()
        {
            return _file.Name.Split('.')[0];
        }


        /// <summary>
        /// Queries the server to determine if a job exists.
        /// </summary>
        /// <exception cref="Exception">Thrown when the request to the server fails</exception>
        /// <returns>True if job exists, false otherwise</returns>
        private async Task<bool> JobExists()
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
                exists = await FabplanHttpRequest.Get(builder.ToString(), new { name = GetJobNumber() });
            }
            catch (Exception e)
            {
                throw new Exception($"Failed to determine if job \"{GetJobNumber()}\" exists.", e);
            }

            return exists;
        }

        /// <summary>
        /// Queries the server to determine if a job is linked to a batch.
        /// </summary>
        /// <exception cref="Exception">Thrown when the request to the server fails</exception>
        /// <returns>True if job is linked to a batch, false otherwise</returns>
        private async Task<bool> JobIsLinked()
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
                isLinked = await FabplanHttpRequest.Get(builder.ToString(), new { name = GetJobNumber() });
            }
            catch (Exception e)
            {
                throw new Exception($"Failed to determine if job \"{GetJobNumber()}\" is linked to a batch.", e);
            }

            return isLinked;
        }
    }

    /// <summary>
    /// An exception thrown when a job cannot be overwritten. 
    /// </summary>
    public sealed class CannotOverwriteJobException : Exception
    {
        public CannotOverwriteJobException(string message, Exception innerException) : base(message, innerException)
        { }
    }

    /// <summary>
    /// An exception thrown when a job cannot be created. 
    /// </summary>
    public sealed class CannotCreateJobException : Exception
    {
        public CannotCreateJobException(string message, Exception innerException) : base(message, innerException)
        { }
    }

    /// <summary>
    /// A class that represents a single entry from a csv job file
    /// </summary>
    class EntreeCSV
    {
        private string _ligne;
        private string _modele;
        private string _essence;
        private string _grade;
	    private string _quantite;
	    private string _hauteur;
	    private string _largeur;
        private string _grain;
        private string _type;
	    private string _profil;
	    private string _dimA;
	    private string _dimB;
	    private string _dimB1;
	    private string _dimC;
	    private string _dimH1;
	    private string _dimH2;
        private string _dueDate;

        
        /// <summary>
        /// Main constructor that accepts line from a job csv file as a parameter 
        /// </summary>
        /// <param name="ligne"></param>
        public EntreeCSV(string ligne)
        {
		    Ligne = ligne;
		    Parse();
        }

        public string Ligne { get => _ligne; set => _ligne = value; }
        public string Modele { get => _modele; set => _modele = value; }
        public string Essence { get => _essence; set => _essence = value; }
        public string Grade { get => _grade; set => _grade = value; }
        public string Quantite { get => _quantite; set => _quantite = value; }
        public string Hauteur { get => _hauteur; set => _hauteur = value; }
        public string Largeur { get => _largeur; set => _largeur = value; }
        public string Grain { get => _grain; set => _grain = value; }
        public string Type { get => _type; set => _type = value; }
        public string Profil { get => _profil; set => _profil = value; }
        public string DimA { get => _dimA; set => _dimA = value; }
        public string DimB { get => _dimB; set => _dimB = value; }
        public string DimB1 { get => _dimB1; set => _dimB1 = value; }
        public string DimC { get => _dimC; set => _dimC = value; }
        public string DimH1 { get => _dimH1; set => _dimH1 = value; }
        public string DimH2 { get => _dimH2; set => _dimH2 = value; }
        public string DueDate { get => _dueDate; set => _dueDate = value; }

        /// <summary>
        /// Parses the line fro the csv file and extracts its data. 
        /// </summary>
        private void Parse()
        {
            List<string> split = Ligne.Split(';').ToList();

            while (split.Count < 16)
            {
                split.Add("");
            }

            Modele = split[0];
            Essence = split[1];
            Grade = split[2];
		    Quantite = split[3];
		    Hauteur = split[4];
		    Largeur = split[5];
		    Grain = split[6];
            Type = split[7];
            Profil = split[8];
		    DimA = split[9];
		    DimB = split[10];
		    DimB1 = split[11];
		    DimC = split[12];
		    DimH1 = split[13];
		    DimH2 = split[14];
            DueDate = split[15];

            if (Type.Trim() == "")
            {
                Type = "0";
            }
            else
            {
                Type = new Regex("^([0-9]+)[\\.]?[0-9]*$").Replace(Type, "$1");
            }
        }
    }

}