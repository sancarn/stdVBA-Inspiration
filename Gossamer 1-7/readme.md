Src: https://www.vbforums.com/showthread.php?532752-VB6-Yet-Another-Web-Server
Author: dilettante 

# Gossamer 1.7

Version 1.0  created 22 Jul 2008
Version 1.7  updated 18 Oct 2013

A simple HTTP server control for use in VB6 projects.

Gossamer can be dropped onto your main Form to provide your application
with an embedded web server.  Gossamer will serve static content from a
directory you specify, or can optionally raise events back to its
parent container allowing you to process dynamic requests.

Currently Gossamer relies on a GET request with parameters or a POST
request to decide a request is dynamic.  One might enhance this to
accept a set of resource extension values or directories to signal a
dynamic request instead (or in addition).


# Structure:

    ------------------------------------------
    | Gossamer                               |

    |                                        |

    | -------------------------------------  |

    | | Winsock control: wskRequest       |  |

    | -------------------------------------  |

    | -------------------------------------- |

    | | Control array of GossClient        | |

    | |                                    | |

    | | ---------------------------------- | |

    | | | GossClient(0)                  | | |

    | | |                                | | |

    | | | ------------------------------ | | |

    | | | | Winsock control: wskClient | | | |

    | | | ------------------------------ | | |

    | | ---------------------------------- | |

    | | ---------------------------------- | |

    | | | GossClient(1)                  | | |

    | | |                                | | |

    | | ---------------------------------- | |

    | |                 :                  | |
    | |                 :                  | |

    | -------------------------------------- |

    ------------------------------------------


# Add Gossamer:

To add Gossamer to your Project you copy the Gossamer file to your
Project folder:

    Gossamer.ctl
    Gossamer.ctx
    GossClient.ctl
    GossClient.ctx
    GossEvent.cls

Then from with the VB6 IDE you can add the two UserControls and the
Class to the project.

To use Gossamer on a Form just add an instance of Gossamer via the IDE
toobox.  You can set Gossamer's properties in design mode or at
runtime.  GossClient is only used internally by Gossamer.

GossEvent objects are used to pass loggable events to your program.


# Using Gossamer:

The sample Project "GossDemo1" shows the basic use of Gossamer.  This
example illustrates how to process GET requests with parameters and
POST requests in a simple manner, as well as logging the server
events supported by Gossamer.

If you only want to serve static content simply don't handle Gossamer
DynamicRequest events.  Gossamer will return "501 Not Implemented"
responses.
