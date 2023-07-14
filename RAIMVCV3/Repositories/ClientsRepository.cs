using RAIMVCV3.Models;
using System;
using System.Collections.Generic;
using System.Data.Entity;
using System.Diagnostics;
using System.Linq;
using System.Web;

namespace RAIMVCV3.Repository
{
    public class ClientsRepository
    {
        public List<Client> GetClients()
        {
            using (ApplicationDbContext context = GetContext())
            {
                return context.Client
                    .OrderBy(c => c.ClientName)
                  .ToList();
            }
        }
        public Client GetClient(int id)
        {
            using (ApplicationDbContext context = GetContext())
            {
                return context.Client                 
                   .Where(cb => cb.ClientID == id)
                   .SingleOrDefault();
            }
        }

        public void AddClient(Client client)
        {
            using (ApplicationDbContext context = GetContext())
            {
                context.Client.Add(client);

                context.SaveChanges();
            }
        }
        public void UpdateClient(Client client)
        {
            using (ApplicationDbContext context = GetContext())
            {
                context.Client.Attach(client);

                var clientEntry = context.Entry(client);
                clientEntry.State = EntityState.Modified;
                //comicBookEntry.Property("IssueNumber").IsModified = false;

                context.SaveChanges();
            }
        }
        public void DeleteClient(int clientId)
        {
            using (ApplicationDbContext context = GetContext())
            {
                var client = new Client() { ClientID = clientId };
                context.Entry(client).State = EntityState.Deleted;

                context.SaveChanges();
            }
        }
        static ApplicationDbContext GetContext()
        {
            var context = new ApplicationDbContext();
            context.Database.Log = (message) => Debug.WriteLine(message);
            return context;
        }
    }
}