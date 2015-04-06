using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace O365_Win_Snippets
{

    public class StoryDefinition : ViewModelBase
    {
        public string GroupName { get; set; }
        public string Title { get; set; }

        // Delegate method to call
        public Func<Task<bool>> RunStoryAsync { get; set; }

        bool _isRunning = false;
        public bool IsRunning
        {
            get
            {
                return _isRunning;
            }
            set
            {
                SetProperty(ref _isRunning, value);
            }
        }

        bool? _result = null;
        public bool? Result
        {
            get
            {
                return _result;
            }
            set
            {
                SetProperty(ref _result, value);
            }
        }

        long? _durationMS = 0;
        public long? DurationMS
        {
            get
            {
                return _durationMS;
            }
            set
            {
                SetProperty(ref _durationMS, value);
            }
        }

    }

    public class TestGroup
    {
        public string GroupTitle { get; set; }
        public List<StoryDefinition> Tests { get; set; }
    }


}
