using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace StudyGuide.Logic.Boundaries
{
    public interface IMnemonicService
    {
        string GetMnemonic(string abbreviation);
    }
}
