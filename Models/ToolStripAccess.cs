using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Pinnacle
{
    public interface ToolStripAccess
    {
        void News();
        void Saves();
        void Prints();
        void Searchs();
        void Searchs(int EditID);
        void Deletes();
        void ReadOnlys();
        void Imports();
        void Pdfs();
        void ChangePasswords();
        void DownLoads();
        void ChangeSkins();
        void Logins();
        void GlobalSearchs();
        void TreeButtons();
        void Exit();
        void GridLoad();
       

    }
}
