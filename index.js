import {PythonShell} from 'python-shell';

PythonShell.run('script.py', null, function (err, results) {
  if (err) throw err;
  // results is an array consisting of messages collected during execution
  console.log('results: %j', results);
});