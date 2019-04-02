using System;
using System.Collections.Generic;
using System.Linq;

//I chose "_" prefix to show input parameters
//I chose camelCase

namespace Template.Invalidate
{
    public class Invalidation
    {
        public void SetField<T>(ref T _field, T _value)
        {
            ValidationContainer.TryAdd(_field.GetHashCode(), (DefaultFunc, null));

            if (!EqualityComparer<T>.Default.Equals(_field, _value))
            {
                Target = _field = (T)IsInvalid(_value);
            }
        }
        public object Target { get; set; }
        
        public Func<object, object> DefaultFunc => IsInvalid;

        Dictionary<object, (Func<object, object> Func, object[] Parameters)> ValidationContainer = new Dictionary<object, (Func<object, object> Func, object[] Parameters)>(); 
        public bool AddInvalidation(Func<object, object> _function, object[] _parameters)
        {
            if (_function == null)
            {
                return false;
            }
            if(_parameters == null)
            {
                return false;
            }
            if (Target != null)
            {
                if(ValidationContainer.TryAdd(ValidationContainer.Count, (_function, _parameters)))
                {
                    return true;
                }
            }
            return false;
        }

        //Recursivly apply all functions to target object
        public object IsInvalid(object _input)
        {
            bool invalid = false;
            //check if there are validations to run, if not return the input.
            if (ValidationContainer.Count > 1)
            {
                return ApplyValidations(_input, 1, ref invalid);
            }
            return _input;
        }

        private object ApplyValidations(object _input, int _iterationResult, ref bool _invalid)
        {
            int _recursor = _iterationResult;
            //number of functions to run
            int count = ValidationContainer.Count();

            //Iterator is less then total number of functions.
            if (_recursor < count)
            {
                Func<object, object> Method = ValidationContainer[_recursor].Func;
                try
                {
                    //call the function and check to see if the return value is the same. 
                    if (Method(_input) != null)
                    {
                        //increase iterator result
                        _recursor += 1;
                        //Apply Next Validation
                        ApplyValidations(_input, _recursor, ref _invalid);
                    }
                    else
                    {
                        //Mark it invalid
                        _invalid = true;
                    }
                }
                catch(Exception ex)
                {
                    while(ex.InnerException != null)
                    {
                        Console.WriteLine(ex.Message);
                        ex = ex.InnerException;
                    }
                    Console.WriteLine(ex.Message);
                }

            }
            //if Invalid return _input or new object();
            return _invalid == false ? _input : new object();
        }
    }
}