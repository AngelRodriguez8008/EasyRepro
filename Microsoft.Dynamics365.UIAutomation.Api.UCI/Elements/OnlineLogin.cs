﻿// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT license.
using System;
using System.Security;
using Microsoft.Dynamics365.UIAutomation.Browser;

namespace Microsoft.Dynamics365.UIAutomation.Api.UCI
{
    public class OnlineLogin : Element
    {
        private WebClient _client;

        public OnlineLogin(WebClient client) 
        {
            _client = client;
        }

        /// <summary>
        /// Logs into the organization with the user and password provided
        /// </summary>
        /// <param name="orgUrl">URL of the organization</param>
        /// <param name="username">User name</param>
        /// <param name="password">Password</param>
        /// <param name="mfaSecrectKey">SecrectKey for multi-factor authentication</param>
        public void Login(Uri orgUrl, SecureString username, SecureString password, SecureString mfaSecrectKey = null)
        {
            _client.Login(orgUrl, username, password, mfaSecrectKey);

            if (_client.Browser.Options.UCITestMode)
            {
                _client.InitializeTestMode(true);
            }
        }

        /// <summary>
        /// Logs into the organization with the user and password provided
        /// </summary>
        /// <param name="orgUrl">URL of the organization</param>
        /// <param name="username">User name</param>
        /// <param name="password">Password</param>
        /// <param name="mfaSecrectKey">SecrectKey for multi-factor authentication</param>
        /// <param name="redirectAction">Actions required during redirect</param>
        public void Login(Uri orgUrl, SecureString username, SecureString password, SecureString mfaSecrectKey, Action<LoginRedirectEventArgs> redirectAction)
        {
            _client.Login(orgUrl, username, password, mfaSecrectKey, redirectAction);

            if (_client.Browser.Options.UCITestMode)
            {
                _client.InitializeTestMode(true);
            }
        }

    }
}
