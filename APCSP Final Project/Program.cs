using System.Collections.Generic;

namespace APCSP_Final_Project
{
    class Program
    {
        static void Main(string[] args)
        {
            var db = new DatabaseAccess();
            long cardTotal = db.CountCards();
            List<Card> data = db.GetCards(cardTotal);
            var clerk = new Clerk();
            clerk.connect();
            clerk.createService("APCSP Final Project");
            var id = clerk.createSheet();
            clerk.writeSheet(id, data);
        }
    }
}