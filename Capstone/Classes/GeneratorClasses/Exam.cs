using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace Capstone.Classes.GeneratorClasses
{
    public class Exam
    {
        private string code, name, section, faculty, proctor, room, day, start, end, dur;

        public string Code { get => code; set => code = value; }
        public string Name { get => name; set => name = value; }
        public string Section { get => section; set => section = value; }
        public string Faculty { get => faculty; set => faculty = value; }
        public string Proctor { get => proctor; set => proctor = value; }
        public string Room { get => room; set => room = value; }
        public string Day { get => day; set => day = value; }
        public string Start { get => start; set => start = value; }
        public string End { get => end; set => end = value; }
        public string Duration { get => dur; set => dur = value; }
    }
}