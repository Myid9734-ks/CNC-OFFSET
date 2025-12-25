using System;
using System.Collections.Generic;
using System.Linq;

namespace FanucFocasTutorial
{
    public class CNCConnectionManager : IDisposable
    {
        private readonly Dictionary<string, CNCConnection> _connections;
        private bool _disposed;

        public CNCConnectionManager()
        {
            _connections = new Dictionary<string, CNCConnection>();
            _disposed = false;
        }

        public bool AddConnection(string ipAddress, ushort port = 8193, int timeout = 2)
        {
            if (_connections.ContainsKey(ipAddress))
            {
                return false;
            }

            var connection = new CNCConnection(ipAddress, port, timeout);
            _connections.Add(ipAddress, connection);
            return true;
        }

        public bool RemoveConnection(string ipAddress)
        {
            if (_connections.TryGetValue(ipAddress, out var connection))
            {
                connection.Dispose();
                return _connections.Remove(ipAddress);
            }
            return false;
        }

        public bool ConnectAll()
        {
            bool allConnected = true;
            foreach (var connection in _connections.Values)
            {
                if (!connection.Connect())
                {
                    allConnected = false;
                }
            }
            return allConnected;
        }

        public void DisconnectAll()
        {
            foreach (var connection in _connections.Values)
            {
                connection.Disconnect();
            }
        }

        public CNCConnection GetConnection(string ipAddress)
        {
            return _connections.TryGetValue(ipAddress, out var connection) ? connection : null;
        }

        public IEnumerable<CNCConnection> GetAllConnections()
        {
            return _connections.Values;
        }

        public void Dispose()
        {
            Dispose(true);
            GC.SuppressFinalize(this);
        }

        protected virtual void Dispose(bool disposing)
        {
            if (!_disposed)
            {
                if (disposing)
                {
                    foreach (var connection in _connections.Values)
                    {
                        connection.Dispose();
                    }
                    _connections.Clear();
                }
                _disposed = true;
            }
        }

        ~CNCConnectionManager()
        {
            Dispose(false);
        }
    }
} 