#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Created on Mon Mar  4 17:36:36 2019

@author: bahram.vazir.nezhad@accenture.com
"""

import dialogflow_v2 as dialogflow

class intentManage ():
    
    def list_intents(project_id):
        intents_client = dialogflow.IntentsClient()
    
        parent = intents_client.project_agent_path(project_id)
#        print (parent)
        intents = intents_client.list_intents(parent)
        intentIds =[]
        for intent in intents:
            intentIds.append (format (intent.name))
            
#        for intent in intents:
#            print('=' * 20)
#            print('Intent name: {}'.format(intent.name))
#            print('Intent display_name: {}'.format(intent.display_name))
#            print('Action: {}\n'.format(intent.action))
#            print('Root followup intent: {}'.format(
#                intent.root_followup_intent_name))
#            print('Parent followup intent: {}\n'.format(
#                intent.parent_followup_intent_name))
#    
#            print('Input contexts:')
#            for input_context_name in intent.input_context_names:
#                print('\tName: {}'.format(input_context_name))
#    
#            print('Output contexts:')
#            for output_context in intent.output_contexts:
#                print('\tName: {}'.format(output_context.name))
        return intentIds

    def detect_intent_texts(project_id, session_id, texts, language_code):
        """Returns the result of detect intent with texts as inputs.
    
        Using the same `session_id` between requests allows continuation
        of the conversation."""
    
        import dialogflow_v2 as dialogflow
        session_client = dialogflow.SessionsClient()
    
        session = session_client.session_path(project_id, session_id)
#        print('Session path: {}\n'.format(session))
    
        for text in texts:
            text_input = dialogflow.types.TextInput(
                text=text, language_code=language_code)
    
            query_input = dialogflow.types.QueryInput(text=text_input)
    
            response = session_client.detect_intent(
                session=session, query_input=query_input)
    
#            print('=' * 20)
#            print('Query text: {}'.format(response.query_result.query_text))
            print('Detected intent: {} (confidence: {})\n'.format(
                response.query_result.intent.display_name,
                response.query_result.intent_detection_confidence))
#            print('Fulfillment text: {}\n'.format(
#                response.query_result.fulfillment_text))
            return [response.query_result.intent.display_name, response.query_result.intent_detection_confidence, response.query_result.fulfillment_text]
            
    def create_intent(project_id, display_name, training_phrases_parts,
                      message_texts):
        """Create an intent of the given intent type."""
        intents_client = dialogflow.IntentsClient()
        parent = intents_client.project_agent_path(project_id)
        training_phrases = []
        for training_phrases_part in training_phrases_parts:
            part = dialogflow.types.Intent.TrainingPhrase.Part(
                text=training_phrases_part)
            # Here we create a new training phrase for each provided part.
            training_phrase = dialogflow.types.Intent.TrainingPhrase(parts=[part])
            training_phrases.append(training_phrase)
        text = dialogflow.types.Intent.Message.Text(text=message_texts)
        message = dialogflow.types.Intent.Message(text=text)
        intent = dialogflow.types.Intent(
            display_name=display_name,
            training_phrases=training_phrases,
            messages=[message])
        response = intents_client.create_intent(parent, intent)
        print('Intent created: {}'.format(response))
    
    def delete_intent(project_id, intent_id):
        """Delete intent with the given intent type and intent value."""
        intents_client = dialogflow.IntentsClient()
    
        intent_path = intents_client.intent_path(project_id, intent_id)
    
        intents_client.delete_intent(intent_path)

    def detect_intent_audio(project_id, session_id, audio_file_path,
                            language_code):
        """Returns the result of detect intent with an audio file as input.
    
        Using the same `session_id` between requests allows continuation
        of the conversaion."""
        import dialogflow_v2 as dialogflow
    
        session_client = dialogflow.SessionsClient()
    
        # Note: hard coding audio_encoding and sample_rate_hertz for simplicity.
        audio_encoding = dialogflow.enums.AudioEncoding.AUDIO_ENCODING_LINEAR_16
        sample_rate_hertz = 16000
    
        session = session_client.session_path(project_id, session_id)
        print('Session path: {}\n'.format(session))
    
        with open(audio_file_path, 'rb') as audio_file:
            input_audio = audio_file.read()
    
        audio_config = dialogflow.types.InputAudioConfig(
            audio_encoding=audio_encoding, language_code=language_code,
            sample_rate_hertz=sample_rate_hertz)
        query_input = dialogflow.types.QueryInput(audio_config=audio_config)
    
        response = session_client.detect_intent(
            session=session, query_input=query_input,
            input_audio=input_audio)
    
        print('=' * 20)
        print('Query text: {}'.format(response.query_result.query_text))
        print('Detected intent: {} (confidence: {})\n'.format(
            response.query_result.intent.display_name,
            response.query_result.intent_detection_confidence))
        print('Fulfillment text: {}\n'.format(
            response.query_result.fulfillment_text))
        
    def detect_intent_stream(project_id, session_id, audio_file_path,
                             language_code):
        """Returns the result of detect intent with streaming audio as input.
    
        Using the same `session_id` between requests allows continuation
        of the conversaion."""
        import dialogflow_v2 as dialogflow
        session_client = dialogflow.SessionsClient()
    
        # Note: hard coding audio_encoding and sample_rate_hertz for simplicity.
        audio_encoding = dialogflow.enums.AudioEncoding.AUDIO_ENCODING_LINEAR_16
        sample_rate_hertz = 16000
    
        session_path = session_client.session_path(project_id, session_id)
        print('Session path: {}\n'.format(session_path))
    
        def request_generator(audio_config, audio_file_path):
            query_input = dialogflow.types.QueryInput(audio_config=audio_config)
    
            # The first request contains the configuration.
            yield dialogflow.types.StreamingDetectIntentRequest(
                session=session_path, query_input=query_input)
    
            # Here we are reading small chunks of audio data from a local
            # audio file.  In practice these chunks should come from
            # an audio input device.
            with open(audio_file_path, 'rb') as audio_file:
                while True:
                    chunk = audio_file.read(4096)
                    if not chunk:
                        break
                    # The later requests contains audio data.
                    yield dialogflow.types.StreamingDetectIntentRequest(
                        input_audio=chunk)
    
        audio_config = dialogflow.types.InputAudioConfig(
            audio_encoding=audio_encoding, language_code=language_code,
            sample_rate_hertz=sample_rate_hertz)
    
        requests = request_generator(audio_config, audio_file_path)
        responses = session_client.streaming_detect_intent(requests)
    
        print('=' * 20)
        for response in responses:
            print('Intermediate transcript: "{}".'.format(
                    response.recognition_result.transcript))
    
        # Note: The result from the last response is the final transcript along
        # with the detected content.
        query_result = response.query_result
    
        print('=' * 20)
        print('Query text: {}'.format(query_result.query_text))
        print('Detected intent: {} (confidence: {})\n'.format(
            query_result.intent.display_name,
            query_result.intent_detection_confidence))
        print('Fulfillment text: {}\n'.format(
            query_result.fulfillment_text))

    def create_entity_type(project_id, display_name, kind):
        """Create an entity type with the given display name."""
        import dialogflow_v2 as dialogflow
        entity_types_client = dialogflow.EntityTypesClient()
    
        parent = entity_types_client.project_agent_path(project_id)
        entity_type = dialogflow.types.EntityType(
            display_name=display_name, kind=kind)
    
        response = entity_types_client.create_entity_type(parent, entity_type)
    
        print('Entity type created: \n{}'.format(response))
    
    def delete_entity_type(project_id, entity_type_id):
        """Delete entity type with the given entity type name."""
        import dialogflow_v2 as dialogflow
        entity_types_client = dialogflow.EntityTypesClient()
    
        entity_type_path = entity_types_client.entity_type_path(
            project_id, entity_type_id)
    
        entity_types_client.delete_entity_type(entity_type_path)
    
    def create_entity(project_id, entity_type_id, entity_value, synonyms):
        """Create an entity of the given entity type."""
        import dialogflow_v2 as dialogflow
        entity_types_client = dialogflow.EntityTypesClient()
    
        # Note: synonyms must be exactly [entity_value] if the
        # entity_type's kind is KIND_LIST
        synonyms = synonyms or [entity_value]
    
        entity_type_path = entity_types_client.entity_type_path(
            project_id, entity_type_id)
    
        entity = dialogflow.types.EntityType.Entity()
        entity.value = entity_value
        entity.synonyms.extend(synonyms)
    
        response = entity_types_client.batch_create_entities(
            entity_type_path, [entity])
    
        print('Entity created: {}'.format(response))
    
    def delete_entity(project_id, entity_type_id, entity_value):
        """Delete entity with the given entity type and entity value."""
        import dialogflow_v2 as dialogflow
        entity_types_client = dialogflow.EntityTypesClient()
    
        entity_type_path = entity_types_client.entity_type_path(
            project_id, entity_type_id)
    
        entity_types_client.batch_delete_entities(
            entity_type_path, [entity_value])
    
    def create_session_entity_type(project_id, session_id, entity_values,
                                   entity_type_display_name, entity_override_mode):
        """Create a session entity type with the given display name."""
        import dialogflow_v2 as dialogflow
        session_entity_types_client = dialogflow.SessionEntityTypesClient()
    
        session_path = session_entity_types_client.session_path(
            project_id, session_id)
        session_entity_type_name = (
            session_entity_types_client.session_entity_type_path(
                project_id, session_id, entity_type_display_name))
    
        # Here we use the entity value as the only synonym.
        entities = [
            dialogflow.types.EntityType.Entity(value=value, synonyms=[value])
            for value in entity_values]
        session_entity_type = dialogflow.types.SessionEntityType(
            name=session_entity_type_name,
            entity_override_mode=entity_override_mode,
            entities=entities)
    
        response = session_entity_types_client.create_session_entity_type(
            session_path, session_entity_type)
    
        print('SessionEntityType created: \n\n{}'.format(response))
    
    def delete_session_entity_type(project_id, session_id,
                                   entity_type_display_name):
        """Delete session entity type with the given entity type display name."""
        import dialogflow_v2 as dialogflow
        session_entity_types_client = dialogflow.SessionEntityTypesClient()
    
        session_entity_type_name = (
            session_entity_types_client.session_entity_type_path(
                project_id, session_id, entity_type_display_name))
    
        session_entity_types_client.delete_session_entity_type(
            session_entity_type_name)
