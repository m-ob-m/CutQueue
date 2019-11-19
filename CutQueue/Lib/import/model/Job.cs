using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Linq;

namespace CutQueue.Lib.Import.Model
{
    /// <summary>
    /// A class that represents a job.
    /// </summary>
    public class Job
    {
        public string Identifier { get; set; } = null;
        public string PONumber { get; set; } = null;
        public DateTime? RequiredDate { get; set; } = null;
        public Customer Customer { get; set; } = new Customer();
        public Material Material { get; set; } = new Material();
        [JsonConverter(typeof(JobTypeListJsonConverter))]
        public SortedList<int, JobType> JobTypeList { get; set; } = new SortedList<int, JobType>();
    }

    /// <summary>
    /// A class that represents a model-type section in a job.
    /// </summary>
    public class JobType
    {
        public string Model { get; set; } = null;
        public int? Type { get; set; } = null;
        public string ExternalProfile { get; set; } = null;
        public List<Part> Parts { get; set; } = new List<Part>();
    }

    /// <summary>
    /// A class that represents a part in a model-type section in a job.
    /// </summary>
    public class Part
    {
        public int? Quantity { get; set; } = null;
        public decimal? Height { get; set; } = null;
        public decimal? Width { get; set; } = null;
        public string GrainDirection { get; set; } = null;
        /* 0 = No grain direction, 1 = Horizontal grain direction, 2 = vertical Grain direction */
    }

    /// <summary>
    /// A class that represents a model-type section in a job.
    /// </summary>
    public class Customer
    {
        public string Name { get; set; } = null;
        public string Address1 { get; set; } = null;
        public string Address2 { get; set; } = null;
        public string PostalCode { get; set; } = null;
    }

    /// <summary>
    /// A class that represents a material.
    /// </summary>
    public class Material
    {
        public string Essence { get; set; } = null;
        public string Grade { get; set; } = null;
    }

    internal class JobTypeListJsonConverter : JsonConverter
    {
        public override void WriteJson(JsonWriter writer, object value, JsonSerializer serializer)
        {
            serializer.Serialize(writer, (value as SortedList<int, JobType>).Values);
        }

        public override object ReadJson(JsonReader reader, Type objectType, object existingValue, JsonSerializer serializer)
        {
            throw new NotImplementedException("Unnecessary because CanRead is false. The type will skip the converter.");
        }

        public override bool CanRead
        {
            get { return false; }
        }

        public override bool CanConvert(Type objectType)
        {
            return objectType == typeof(SortedList<int, JobType>);
        }
    }
}
