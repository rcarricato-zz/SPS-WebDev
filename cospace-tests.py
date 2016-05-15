import acano
import time

if __name__ == "__main__":
    # Setup acano module to work with test server
    acano.HOST = 'https://172.31.189.189/api/v1'
    acano.API_USER = 'python'
    acano.API_PASS = 'python'
    acano.SSL_VERIFY = False
    
    # In the final version, we'll have some true mechanism for getting a list of
    # users or cospaces, but for now we'll provide test ones
    users = [
        # name          uri                secondary  callId     passcode
        ('pyaccount1', 'join.pyaccount1', '5551001', '5551001', '4121'),
        ('pyaccount2', 'join.pyaccount2', '5551002', '5551002', '4222'),
        ('pyaccount3', 'join.pyaccount3', '5551003', '5551003', '4323'),
    ]

    # Loop over all users in our record set
    for user in users:
        # Figure out if we need to create the coSpace
        cospaces = acano.cospaces_get()
        cospace = None
        for co in cospaces:
            if co['name'] == user[0]:
                print("coSpace %s already exists" % user[0])
                cospace = co
                break
        if not cospace:
            print("Creating new coSpace %s" % user[0])
            new_space = {
                'name': user[0],
                'uri': user[1],
                'secondaryUri': user[2],
                'callId': user[3],
            }
            acano.cospace_create(new_space)
            cospace = acano.cospace_get(new_space['name'])
        if cospace:
            print("coSpace id = %s" % cospace['id'])
        else:
            raise Exception("coSpace creation failed")

        # Create host access for endpoints
        acano.accessmethod_create(cospace, {
            'uri': '%s.%s' % (cospace['callId'], user[4]),
        })

        # Create host access for WebRTC and Audio
        acano.accessmethod_create(cospace, {
            'callId': cospace['callId'],
            'passcode': user[4],
        })

    # Clean up all created spaces
    for user in users:
        print("About to delete coSpace %s..." % user[0])
        time.sleep(2)
        cospace = acano.cospace_get(user[0])
        acano.cospace_delete(cospace)
    print("All done!")
