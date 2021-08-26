// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT license.

using GraphSDKDemo.Models;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;

namespace GraphTutorial.Models
{
    public class MailViewModel
    {

        private ObservableCollection<GraphSDKDemo.Models.Message> _events;

        public MailViewModel(ObservableCollection<GraphSDKDemo.Models.Message> msgs)
        {
            _events = msgs;
        }

        public ObservableCollection<Message> Events { get => _events; set => _events = value; }
    }
}

