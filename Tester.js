function test_PropStore() {
  Logger.log('starting test');
  var baz = 'one two three';
  var foo = NSPropertyHelper.get('test-key');
  NSPropertyHelper.set('test-key', baz);
  var bar = NSPropertyHelper.get('test-key');
  NSPropertyHelper.delete('test-key');
  var result = baz === bar ? "Pass" : "Fail";
  Logger.log(`Test result ${result} -- ${bar} # ${'baz'}`);
}
