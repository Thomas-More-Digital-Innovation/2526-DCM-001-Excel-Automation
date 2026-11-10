// "Suggest Mapping (OpenAI)": only runs when some headers were never seen before.
return [{ json: await suggestUnknown($env, $input.first().json) }];
