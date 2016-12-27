using System;

namespace OfflineConflattionJudgeTool
{
    public class Job
    {
        public Job(Guid id,string name,string createtime,string path)
        {
            JobId = id;
            JobName = name;
            CreateTime = createtime;
            JobPath = path;
            CurrentIndex = 0;
            
        }
        public Guid JobId { get; set; }
        public string JobName { get; set; }
        public string Owner { get; set; }
        public string CreateTime { get; set; }
        public string FinishTime { get; set; }
        public string JobPath { get; set; }
        public int CurrentIndex { get; set; }
        public int Count { get; set; }
        public int HasFinished { get; set; }
    }
}
